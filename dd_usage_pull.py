#!/usr/bin/env python3
"""
dd_usage_pull.py — Datadog → Coralogix Usage Report
=====================================================
Pulls Datadog usage metrics from the official Usage Metering API, mirrors
every tile on the Bill Overview page, converts the numbers to Coralogix
sizing using the same formulas as the sizing Excel template, then packages
all outputs into a single ZIP file ready to hand off.

Usage
-----
    python dd_usage_pull.py [--month YYYY-MM] [--site datadoghq.com] [--out DIR]

Environment variables (put these in a .env file next to the script):
    DD_API_KEY   Datadog API key          (needs usage_read)
    DD_APP_KEY   Datadog Application key  (needs billing_read for cost data)
    DD_SITE      datadoghq.com | datadoghq.eu | us3/us5.datadoghq.com  (default: datadoghq.com)
    DD_MONTH     YYYY-MM  (default: previous calendar month)
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import sys
import time
import zipfile
from dataclasses import dataclass, field, fields as dc_fields
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

# ── Dependency guards ──────────────────────────────────────────────────────
def _import(mod: str, pkg: str) -> Any:
    try:
        import importlib
        return importlib.import_module(mod)
    except ImportError:
        sys.exit(f"\n  Missing '{mod}'. Run:  pip install {pkg}\n")

requests_mod = _import("requests", "requests")
dotenv_mod   = _import("dotenv",   "python-dotenv")

import requests                           # noqa: E402  (already verified above)
from dotenv import load_dotenv            # noqa: E402

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("  [warn] openpyxl not installed — Excel output will be skipped. Run: pip install openpyxl")

load_dotenv()

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 1 — Constants  (all sizing assumptions live here)
# ═══════════════════════════════════════════════════════════════════════════

AVG_LOG_SIZE_KB  = 2.5     # assumed average compressed log line size
AVG_SPAN_SIZE_KB = 1.5     # assumed average span payload size
TS_PER_HOST      = 750     # estimated Prometheus/DD time series per host or container
TS_TO_UNITS      = 3.3e-5  # Coralogix metrics: TimeSeries → Units/day conversion
DAYS_PER_MONTH   = 30.0    # used for all monthly → daily conversions
HOURS_PER_MONTH  = 30.0 * 24  # 720 — Datadog _sum fields for hosts/containers are in host-hours
SW_LABEL_FACTOR  = 0.30    # serverless TS = invocations × 3 labels × 10% unique = ×0.30
LOG_TIER_MON     = 0.70    # Monitoring share of ingested logs
LOG_TIER_COMP    = 0.30    # Compliance share

KNOWN_SITES: dict[str, str] = {
    "datadoghq.com":     "https://api.datadoghq.com",
    "datadoghq.eu":      "https://api.datadoghq.eu",
    "us3.datadoghq.com": "https://api.us3.datadoghq.com",
    "us5.datadoghq.com": "https://api.us5.datadoghq.com",
    "ddog-gov.com":      "https://api.ddog-gov.com",
}

FRESHNESS_NOTE = (
    "Datadog usage data may be delayed up to 72 hours. "
    "Historical cost for a closed month becomes available by the 16th of the following month."
)

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 2 — Datadog API client
# ═══════════════════════════════════════════════════════════════════════════

class DatadogClient:
    """Thin, retry-aware wrapper around Datadog's Usage Metering API."""

    MAX_RETRIES   = 4
    RETRY_WAIT_S  = 2  # exponential: 2, 4, 8 seconds

    def __init__(self, api_key: str, app_key: str, site: str = "datadoghq.com"):
        base = KNOWN_SITES.get(site, f"https://api.{site}")
        self.base_url = base.rstrip("/")
        self.headers  = {
            "Accept":             "application/json",
            "DD-API-KEY":         api_key,
            "DD-APPLICATION-KEY": app_key,
        }

    def _get(self, path: str, params: dict | None = None) -> dict:
        url = f"{self.base_url}{path}"
        for attempt in range(self.MAX_RETRIES):
            try:
                resp = requests.get(url, headers=self.headers, params=params or {}, timeout=45)
                if resp.status_code == 429:
                    wait = self.RETRY_WAIT_S ** (attempt + 1)
                    print(f"    Rate limited — waiting {wait}s (retry {attempt+1}/{self.MAX_RETRIES})")
                    time.sleep(wait)
                    continue
                if resp.status_code == 403:
                    raise PermissionError(
                        f"403 Forbidden: {path}\n"
                        "  Ensure the API key has usage_read and the App key has billing_read."
                    )
                if resp.status_code == 400:
                    raise ValueError(f"400 Bad Request: {path} — {resp.text[:300]}")
                resp.raise_for_status()
                return resp.json()
            except (requests.ConnectionError, requests.Timeout) as exc:
                if attempt == self.MAX_RETRIES - 1:
                    raise
                wait = self.RETRY_WAIT_S ** (attempt + 1)
                print(f"    Network error ({exc}) — retrying in {wait}s")
                time.sleep(wait)
        raise RuntimeError(f"All retries exhausted for {path}")

    # ── Individual endpoint wrappers ────────────────────────────────────────

    def usage_summary(self, start_month: str, end_month: str | None = None) -> dict:
        p: dict[str, str] = {"start_month": start_month}
        if end_month:
            p["end_month"] = end_month
        return self._get("/api/v1/usage/summary", p)

    def billable_summary(self, month: str) -> dict:
        return self._get("/api/v1/usage/billable-summary", {"month": month})

    def estimated_cost(self, start_month: str, end_month: str | None = None) -> dict:
        p: dict[str, str] = {"start_month": start_month}
        if end_month:
            p["end_month"] = end_month
        return self._get("/api/v2/usage/estimated_cost", p)

    def historical_cost(self, start_month: str, end_month: str | None = None) -> dict:
        p: dict[str, str] = {"start_month": start_month, "view": "summary"}
        if end_month:
            p["end_month"] = end_month
        return self._get("/api/v2/usage/historical_cost", p)

    def projected_cost(self) -> dict:
        return self._get("/api/v2/usage/projected_cost", {"view": "summary"})

    def billing_dimension_mapping(self, month: str | None = None) -> dict:
        p: dict[str, str] = {}
        if month:
            p["filter[month]"] = f"{month}-01T00:00:00Z"
        return self._get("/api/v2/usage/billing_dimension_mapping", p)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 3 — Data model
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class UsageSnapshot:
    """All Datadog usage metrics for one calendar month."""

    month:        str   = ""
    account_name: str   = ""
    region:       str   = ""
    site:         str   = ""
    pulled_at:    str   = ""

    # ── Infrastructure ──────────────────────────────────────────────────────
    infra_hosts:    float = 0   # Infra Hosts tile (concurrent count)
    apm_hosts:      float = 0   # APM Hosts tile (concurrent count)
    containers:     float = 0   # concurrent container count (for sizing formulas)
    container_hours: float = 0  # Container Hours tile (total host-hours in month)
    network_hosts:  float = 0   # Network Hosts tile
    dbm_hosts:      float = 0   # DBM Hosts tile

    # ── Profiling ───────────────────────────────────────────────────────────
    profiled_hosts:      float = 0
    profiled_containers: float = 0
    profiled_fargate:    float = 0
    apm_fargate:         float = 0
    fargate_tasks:       float = 0

    # ── Custom Metrics ──────────────────────────────────────────────────────
    custom_metrics:          float = 0  # Custom Metrics tile (time series)
    ingested_custom_metrics: float = 0  # Ingested Custom Metrics tile

    # ── Logs ────────────────────────────────────────────────────────────────
    ingested_logs_bytes:      float = 0   # Ingested Logs tile (bytes)
    indexed_logs_3day:        float = 0   # events
    indexed_logs_7day:        float = 0
    indexed_logs_15day:       float = 0
    indexed_logs_30day:       float = 0
    indexed_logs_45day:       float = 0
    indexed_logs_60day:       float = 0
    indexed_logs_90day:       float = 0
    indexed_logs_180day:      float = 0
    indexed_logs_360day:      float = 0
    indexed_logs_live:        float = 0   # live search (short-term)
    indexed_logs_rehydrated:  float = 0   # rehydrated from archive
    security_logs_bytes:      float = 0   # SIEM analyzed logs

    # ── APM / Tracing ───────────────────────────────────────────────────────
    ingested_spans_bytes:  float = 0   # Ingested Spans tile (bytes)
    indexed_spans:         float = 0   # Indexed Spans tile (events)
    custom_events:         float = 0   # Custom Events tile

    # ── Serverless ──────────────────────────────────────────────────────────
    serverless_functions:     float = 0   # Serverless Workload Functions tile
    serverless_invocations:   float = 0   # total invocations (monthly)
    serverless_app_instances: float = 0   # Serverless App Instances tile

    # ── RUM ─────────────────────────────────────────────────────────────────
    rum_sessions:          float = 0   # RUM Investigate tile
    rum_lite_sessions:     float = 0   # RUM Measure tile
    rum_replay:            float = 0   # Session Replay tile
    rum_errors:            float = 0   # Error Tracking Events

    # ── Synthetics ──────────────────────────────────────────────────────────
    synthetics_api:     float = 0
    synthetics_browser: float = 0

    # ── Other ───────────────────────────────────────────────────────────────
    incident_management_seats:     float = 0
    test_optimization_committers:  float = 0
    test_optimization_spans:       float = 0
    product_analytics_sessions:    float = 0
    session_replay:                float = 0   # may alias rum_replay
    app_builder_apps:              float = 0
    bits_ai_investigations:        float = 0


    # ── Full API responses for raw JSON export ───────────────────────────────
    raw: dict = field(default_factory=dict)

    # ── Per-dimension breakdown from billable-summary ────────────────────────
    billable_breakdown: list[dict] = field(default_factory=list)


@dataclass
class CoralogixSizing:
    """Coralogix sizing estimates derived from Datadog usage."""

    # ── Logs ────────────────────────────────────────────────────────────────
    total_ingested_logs_gb_month:  float = 0
    total_indexed_logs_count:      float = 0   # events
    indexed_logs_size_gb_month:    float = 0
    indexed_pct_logs:              float = 0
    daily_logs_gb:                 float = 0
    daily_logs_fs_gb:              float = 0   # unused, kept for compat
    daily_logs_mon_gb:             float = 0   # Monitoring       (40 %)
    daily_logs_comp_gb:            float = 0   # Compliance       (10 %)

    # ── Metrics ─────────────────────────────────────────────────────────────
    host_count:               float = 0
    container_count:          float = 0
    host_container_ts:        float = 0
    sw_func_ts:               float = 0
    sw_invoc_ts:              float = 0
    total_ts:                 float = 0
    metrics_units_per_day:    float = 0

    # ── Tracing ─────────────────────────────────────────────────────────────
    ingested_spans_gb_month:  float = 0
    indexed_spans_gb_month:   float = 0
    indexed_pct_spans:        float = 0
    daily_spans_ingest_gb:    float = 0
    daily_spans_indexed_gb:   float = 0
    daily_spans_archive_gb:   float = 0

    # ── RUM ─────────────────────────────────────────────────────────────────
    rum_sessions_monthly:     float = 0
    rum_errors_per_day:       float = 0


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 4 — Extraction helpers
# ═══════════════════════════════════════════════════════════════════════════

def _f(item: dict, *keys: str) -> float:
    """Return the first non-None numeric value found among `keys`, else 0."""
    for k in keys:
        v = item.get(k)
        if v is not None:
            try:
                return float(v)
            except (TypeError, ValueError):
                continue
    return 0.0


def extract_usage_snapshot(
    summary_raw: dict,
    billable_raw: dict | None,
    month: str,
    site: str,
) -> UsageSnapshot:
    """Build a UsageSnapshot from all available API responses."""

    snap = UsageSnapshot(
        month=month,
        site=site,
        pulled_at=datetime.now(timezone.utc).isoformat(),
    )

    # ── Usage summary ────────────────────────────────────────────────────────
    usage_list = summary_raw.get("usage", [])
    if usage_list:
        # Select the item whose date matches the target month, or use the last item.
        item = usage_list[-1]
        for u in usage_list:
            date_str = str(u.get("date", u.get("start_date", "")))
            if date_str.startswith(month):
                item = u
                break

        # Account name: try top-level first, then orgs[0]
        snap.account_name = item.get("account_name", item.get("org_name", ""))
        if not snap.account_name:
            orgs = item.get("orgs", [])
            if orgs:
                snap.account_name = orgs[0].get("account_name", orgs[0].get("org_name", ""))
        snap.region = item.get("region", "")

        # Infrastructure
        # IMPORTANT: Datadog's v2 API uses "_sum" fields that represent TOTAL HOST-HOURS
        # for the month (not a concurrent count). We divide by HOURS_PER_MONTH (720) to
        # recover the average concurrent host count for sizing purposes.
        # "top99p" fields are already a concurrent count (peak), so they're used as-is.

        def _hosts_from_hours(hours_field: str, *top99p_fields: str) -> float:
            """Try top99p (count) first; fall back to hours ÷ 720."""
            for k in top99p_fields:
                v = _f(item, k)
                if v: return v
            h = _f(item, hours_field)
            return h / HOURS_PER_MONTH if h else 0.0

        # Sum all host types; prefer top99p counts, fall back to hours/720
        snap.infra_hosts = (
            _hosts_from_hours("agent_host_sum",       "agent_host_top99p_sum",   "agent_host_top99p")
            + _hosts_from_hours("aws_host_sum",       "aws_host_top99p_sum",     "aws_host_top99p")
            + _hosts_from_hours("azure_host_sum",     "azure_host_top99p_sum",   "azure_host_top99p")
            + _hosts_from_hours("gcp_host_sum",       "gcp_host_top99p_sum",     "gcp_host_top99p")
            + _hosts_from_hours("vsphere_host_sum",   "vsphere_host_top99p_sum", "vsphere_host_top99p")
            + _hosts_from_hours("alibaba_host_sum",   "alibaba_host_top99p_sum")
            + _hosts_from_hours("heroku_host_sum",    "heroku_host_top99p_sum")
            + _hosts_from_hours("opentelemetry_host_sum", "opentelemetry_host_top99p_sum")
        ) or _hosts_from_hours("infra_hours_sum",     "infra_host_top99p_sum",   "infra_host_top99p")

        snap.apm_hosts = _hosts_from_hours(
            "apm_host_sum",
            "apm_host_top99p_sum", "apm_host_top99p",
        ) or _hosts_from_hours("apm_host_incl_usm_sum", "apm_host_incl_usm_top99p")

        snap.network_hosts = _f(item, "npm_host_top99p", "npm_host_top99p_sum") or \
            _hosts_from_hours("npm_host_sum", "network_device_count_top99p_sum")

        snap.dbm_hosts = _f(item, "dbm_host_top99p", "dbm_host_top99p_sum",
                            "dbm_host_database_instance_top99p")

        # Containers: "_sum" = total container-hours for the month; ÷720 = concurrent count
        raw_container_hours = _f(item, "container_sum", "container_count_avg_sum",
                                 "container_avg_sum")
        snap.container_hours = raw_container_hours   # used for the "Container Hours" tile
        snap.containers = (
            _f(item, "container_count_avg", "container_avg")   # already an average — use as-is
            or (raw_container_hours / HOURS_PER_MONTH if raw_container_hours else 0.0)
        )

        # Profiling / Fargate
        # profiling_host_top99p is already a concurrent count; profiling_container_agent_count_sum is hours
        snap.profiled_hosts = _f(item,
            "profiling_host_top99p",
            "profiling_host_count_top99p_sum", "profiling_host_count_top99p",
            "profiling_uncategorized_host_count_top99p",
        )
        _prof_cont_hours = _f(item, "profiling_container_agent_count_sum")
        snap.profiled_containers = (
            _f(item, "profiling_container_agent_count_avg", "profiling_container_count_avg_sum")
            or (_prof_cont_hours / HOURS_PER_MONTH if _prof_cont_hours else 0.0)
        )
        snap.profiled_fargate = _f(item,
            "avg_profiled_fargate_tasks",
            "avg_profiled_fargate_tasks_hw_max_sum",
            "profiling_aas_count_top99p_sum",
            "fargate_container_profiler_profiling_fargate_avg",
        )
        snap.apm_fargate   = _f(item, "apm_fargate_count_avg_sum", "apm_fargate_count_avg")
        snap.fargate_tasks = _f(item, "fargate_tasks_count_avg_sum",
                                "fargate_tasks_count_avg", "fargate_tasks_count_hwm")

        # Custom Metrics
        snap.custom_metrics          = _f(item, "custom_ts_avg", "custom_ts_avg_sum",
                                          "custom_timeseries_avg_sum")
        snap.ingested_custom_metrics = _f(item, "custom_ingested_timeseries_average_sum",
                                          "ingested_custom_timeseries_average_sum",
                                          "custom_live_ts_avg_sum", "custom_live_ts_avg")

        # Logs — ingested bytes
        # "live_ingested_bytes_sum" is the most common field in newer API responses
        snap.ingested_logs_bytes = _f(item,
            "live_ingested_bytes_sum",           # ← most common in 2025+ responses
            "ingested_events_bytes_sum",
            "ingested_events_bytes_agg_sum",
            "billable_ingested_bytes_agg_sum",
            "logs_live_ingested_bytes_agg_sum",
        )

        # Indexed logs — Datadog uses two naming conventions for retention tiers:
        #   logs_indexed_logs_usage_sum_N_day  (newer, e.g. 2025+)
        #   logs_indexed_Nday_agg_sum          (older)
        snap.indexed_logs_live       = _f(item,
            "logs_indexed_live_index_indexed_sum",
            "logs_live_indexed_logs_usage_sum",
            "logs_live_indexed_count_agg_sum",
            "live_indexed_events_sum",
        )
        snap.indexed_logs_rehydrated = _f(item,
            "logs_rehydrated_indexed_count_agg_sum",
            "rehydrated_indexed_events_sum",
        )
        snap.indexed_logs_3day   = _f(item,
            "logs_indexed_logs_usage_sum_3_day",   # newer naming
            "logs_indexed_3day_agg_sum",
            "logs_indexed_3_day_agg_sum",
        )
        snap.indexed_logs_7day   = _f(item,
            "logs_indexed_logs_usage_sum_7_day",
            "logs_indexed_7day_agg_sum",
            "logs_indexed_7_day_agg_sum",
        )
        snap.indexed_logs_15day  = _f(item,
            "logs_indexed_logs_usage_sum_15_day",
            "logs_indexed_15day_agg_sum",
            "logs_indexed_15_day_agg_sum",
            "logs_indexed_logs_indexed_15day_sum",
        )
        snap.indexed_logs_30day  = _f(item,
            "logs_indexed_logs_usage_sum_30_day",
            "logs_indexed_30day_agg_sum",
            "logs_indexed_30_day_agg_sum",
            "logs_indexed_logs_indexed_30day_sum",
        )
        snap.indexed_logs_45day  = _f(item,
            "logs_indexed_logs_usage_sum_45_day",
            "logs_indexed_45day_agg_sum",
            "logs_indexed_45_day_agg_sum",
            "logs_indexed_logs_indexed_45day_sum",
        )
        snap.indexed_logs_60day  = _f(item,
            "logs_indexed_logs_usage_sum_60_day",
            "logs_indexed_60day_agg_sum",
            "logs_indexed_60_day_agg_sum",
        )
        snap.indexed_logs_90day  = _f(item,
            "logs_indexed_logs_usage_sum_90_day",
            "logs_indexed_90day_agg_sum",
            "logs_indexed_90_day_agg_sum",
            "logs_indexed_logs_indexed_90day_sum",
        )
        snap.indexed_logs_180day = _f(item,
            "logs_indexed_logs_usage_sum_180_day",
            "logs_indexed_180day_agg_sum",
            "logs_indexed_180_day_agg_sum",
        )
        snap.indexed_logs_360day = _f(item,
            "logs_indexed_logs_usage_sum_360_day",
            "logs_indexed_360day_agg_sum",
            "logs_indexed_360_day_agg_sum",
        )
        # SIEM / security logs
        snap.security_logs_bytes = _f(item,
            "twol_ingested_events_bytes_sum",      # "2nd line" (SIEM) ingested bytes
            "twol_ingested_events_bytes_agg_sum",
            "siem_ingested_bytes_agg_sum",
            "siem_analyzed_logs_add_on_count_sum",
        )

        # APM / Tracing
        snap.ingested_spans_bytes = _f(item,
            "twol_ingested_events_bytes_sum",  # Tracing Without Limits — primary ingestion field
            "ingested_spans_bytes_agg_sum",
            "ingested_spans_bytes_sum",
        )
        # apm_ingest_gb_sum only covers overages above the 150 GB/APM-host/month free tier;
        # use it only as a last resort when no byte-level field found, and convert GB → bytes.
        if snap.ingested_spans_bytes == 0:
            _apm_overage_gb = _f(item, "apm_ingest_gb_sum", "apm_ingest_only_gb_sum")
            if _apm_overage_gb:
                snap.ingested_spans_bytes = _apm_overage_gb * 1e9

        snap.indexed_spans = _f(item,
            "trace_search_indexed_events_count_sum",   # newer naming
            "trace_search_indexed_events_count_agg_sum",
            "apm_span_custom_agg_sum",
            "indexed_events_count_sum",
        )
        snap.custom_events = _f(item,
            "custom_events_agg_sum", "custom_events_sum",
        )

        # Serverless
        snap.serverless_functions     = _f(item,
            "serverless_func_count_avg_sum", "serverless_func_avg_sum",
            "serverless_func_count_agg_sum",
        )
        snap.serverless_invocations   = _f(item,
            "lambda_invocations_count_agg_sum",
            "serverless_invocation_count_agg_sum",
            "aws_lambda_invocations_sum",
        )
        snap.serverless_app_instances = _f(item,
            "serverless_apps_total_count_hw_max_sum",
            "serverless_apps_azure_count_hw_max_sum",
        )

        # RUM
        # Total RUM Investigate = all session types summed
        rum_total = _f(item,
            "rum_total_session_count_sum",
            "rum_browser_and_mobile_session_count_sum",
            "rum_session_count_sum",
        )
        rum_lite   = _f(item,
            "rum_lite_session_count_sum",
            "rum_browser_lite_session_count_sum",
            "rum_lite_session_count_agg_sum",
            "rum_browser_lite_session_count_agg_sum",
        )
        rum_replay = _f(item,
            "rum_replay_session_count_sum",
            "rum_browser_replay_session_count_sum",
            "rum_replay_session_count_agg_sum",
            "session_replay_count_agg_sum",
        )
        rum_legacy = _f(item,
            "rum_browser_legacy_session_count_sum",
            "browser_legacy_session_count_sum",
        )
        # If the total session count field is missing, sum the parts
        snap.rum_sessions      = rum_total or (rum_lite + rum_replay + rum_legacy)
        snap.rum_lite_sessions = rum_lite
        snap.rum_replay        = rum_replay
        snap.session_replay    = rum_replay  # alias
        snap.rum_errors        = _f(item,
            "error_tracking_error_events_sum",   # newer naming
            "error_tracking_events_sum",
            "error_tracking_events_agg_sum",
            "total_error_tracking_events_agg_sum",
        )

        # Synthetics
        snap.synthetics_api     = _f(item,
            "synthetics_check_calls_count_agg_sum",
            "synthetics_check_calls_count_sum",
        )
        snap.synthetics_browser = _f(item,
            "synthetics_browser_check_calls_count_sum",   # newer naming
            "synthetics_browser_check_calls_count_agg_sum",
            "browser_check_calls_count_agg_sum",
        )

        # Incident / DBM / CI / Other
        snap.incident_management_seats    = _f(item,
            "incident_management_seats_hwm",
            "incident_management_monthly_active_users_hwm",
            "incident_management_monthly_active_users_hw_max_sum",
            "incident_management_monthly_active_users_hw_max",
        )
        snap.test_optimization_committers = _f(item,
            "ci_visibility_pipeline_committers_hwm",   # newer naming
            "ci_visibility_itsm_committers_hw_max_sum",
            "ci_visibility_committers_hw_max_sum",
        )
        snap.test_optimization_spans      = _f(item,
            "ci_pipeline_indexed_spans_sum",            # newer naming
            "ci_test_indexed_spans_agg_sum",
            "ci_test_indexed_spans_sum",
            "ci_visibility_test_indexed_spans_agg_sum",
        )
        snap.product_analytics_sessions   = _f(item,
            "product_analytics_sum",
            "product_analytics_count_agg_sum",
            "product_analytics_session_count_agg_sum",
        )
        snap.app_builder_apps    = _f(item, "published_app_hwm_sum", "published_app_hw_max_sum")
        snap.bits_ai_investigations = _f(item,
            "bits_ai_investigations_sum",
            "bits_ai_total_conversations_agg_sum",
            "bits_ai_investigations_agg_sum",
            "ai_credits_bits_sre_ai_credits_sum",
        )

    # ── Billable summary overlay ─────────────────────────────────────────────
    if billable_raw:
        dim_map: dict[str, tuple[float, float, float, str]] = {}   # dim → (billable, committed, on_demand, unit)
        breakdown: list[dict] = []
        for entry in billable_raw.get("usage", []):
            dim      = str(entry.get("billing_dimension", "")).lower()
            billable = entry.get("account_billable_usage")
            committed= entry.get("account_committed_usage")
            on_demand= entry.get("account_on_demand_usage")
            unit     = entry.get("usage_unit", "")
            if billable is not None:
                dim_map[dim] = (float(billable), float(committed or 0), float(on_demand or 0), unit)
            breakdown.append({
                "billing_dimension": entry.get("billing_dimension", ""),
                "billable":   billable,
                "committed":  committed,
                "on_demand":  on_demand,
                "unit":       unit,
                "start_date": entry.get("start_date", ""),
                "end_date":   entry.get("end_date", ""),
            })
        snap.billable_breakdown = breakdown

        def bd(*dims: str) -> float:
            for d in dims:
                if d in dim_map:
                    return dim_map[d][0]
            return 0.0

        # Override usage-summary values with cleaner billable-summary values
        _infra = bd("infra_host", "infra_host_sum", "agent_host")
        if _infra: snap.infra_hosts = _infra

        _apm = bd("apm_host", "apm_host_sum")
        if _apm: snap.apm_hosts = _apm

        _cont = bd("container", "container_sum", "container_average", "containers")
        if _cont: snap.containers = _cont

        _cm = bd("custom_timeseries", "custom_metrics", "custom_ts")
        if _cm: snap.custom_metrics = _cm

        # For ingested logs the billable unit may be bytes or GB — trust usage-summary for bytes
        _il = bd("ingested_logs", "ingested_logs_bytes")
        if _il and snap.ingested_logs_bytes == 0:
            snap.ingested_logs_bytes = _il

        _is = bd("ingested_spans", "ingested_spans_bytes")
        if _is and snap.ingested_spans_bytes == 0:
            snap.ingested_spans_bytes = _is

        _rum = bd("rum", "rum_session", "rum_browser_mobile_sessions", "rum_total_session")
        if _rum: snap.rum_sessions = _rum

        _replay = bd("rum_replay", "session_replay", "rum_replay_session_count")
        if _replay:
            snap.rum_replay = _replay
            snap.session_replay = _replay

        _sinv = bd("serverless_invocations", "serverless_invocation", "lambda_invocations")
        if _sinv: snap.serverless_invocations = _sinv

        _sfunc = bd("serverless_functions", "serverless_function", "serverless_func")
        if _sfunc: snap.serverless_functions = _sfunc

        # Indexed logs by retention tier from billable-summary
        _3d  = bd("indexed_logs_3_day",  "indexed_logs_3day",  "logs_indexed_3day")
        _7d  = bd("indexed_logs_7_day",  "indexed_logs_7day",  "logs_indexed_7day")
        _15d = bd("indexed_logs_15_day", "indexed_logs_15day", "logs_indexed_15day")
        _30d = bd("indexed_logs_30_day", "indexed_logs_30day", "logs_indexed_30day")
        _90d = bd("indexed_logs_90_day", "indexed_logs_90day", "logs_indexed_90day")
        if _3d:  snap.indexed_logs_3day  = _3d
        if _7d:  snap.indexed_logs_7day  = _7d
        if _15d: snap.indexed_logs_15day = _15d
        if _30d: snap.indexed_logs_30day = _30d
        if _90d: snap.indexed_logs_90day = _90d

    # ── Cost data ────────────────────────────────────────────────────────────
    # ── Store raw for JSON export ────────────────────────────────────────────
    snap.raw = {
        "usage_summary":    summary_raw,
        "billable_summary": billable_raw or {},
    }

    return snap


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 5 — Coralogix conversion  (matches the Excel template formulas)
# ═══════════════════════════════════════════════════════════════════════════

def compute_coralogix_sizing(snap: UsageSnapshot) -> CoralogixSizing:
    """Apply the sizing formulas from the Excel template to the usage snapshot."""

    cx = CoralogixSizing()

    # ── Logs ─────────────────────────────────────────────────────────────────
    ingested_gb = snap.ingested_logs_bytes / 1e9
    security_gb = snap.security_logs_bytes / 1e9
    cx.total_ingested_logs_gb_month = ingested_gb + security_gb

    # Indexed logs = live/short-term + rehydrated + 15-day + 90-day+ retention
    # Map: Live & Rehydrated = (3day + 7day + live + rehydrated)
    #      15-day tier = indexed_15day
    #      90-day+ tier = 30d + 45d + 60d + 90d + 180d + 360d
    indexed_live_rehydrated = (
        snap.indexed_logs_3day
        + snap.indexed_logs_7day
        + snap.indexed_logs_live
        + snap.indexed_logs_rehydrated
    )
    indexed_15d = snap.indexed_logs_15day
    indexed_long = (
        snap.indexed_logs_30day + snap.indexed_logs_45day
        + snap.indexed_logs_60day + snap.indexed_logs_90day
        + snap.indexed_logs_180day + snap.indexed_logs_360day
    )

    cx.total_indexed_logs_count = indexed_live_rehydrated + indexed_15d + indexed_long

    # Size in bytes: count × avg_log_size_kb × 1024 (bytes per KB)
    # Then convert bytes → GB: ÷ 1024³
    indexed_bytes = cx.total_indexed_logs_count * AVG_LOG_SIZE_KB * 1024
    cx.indexed_logs_size_gb_month = indexed_bytes / (1024 ** 3)

    if cx.total_ingested_logs_gb_month > 0:
        cx.indexed_pct_logs = cx.indexed_logs_size_gb_month / cx.total_ingested_logs_gb_month

    cx.daily_logs_gb      = cx.total_ingested_logs_gb_month / DAYS_PER_MONTH
    cx.daily_logs_mon_gb  = cx.daily_logs_gb * LOG_TIER_MON
    cx.daily_logs_comp_gb = cx.daily_logs_gb * LOG_TIER_COMP

    # ── Metrics ──────────────────────────────────────────────────────────────
    # Hosts = infra + apm + profiled + fargate types (from Excel: A4+A8+B8+F6+C10+B10)
    cx.host_count = (
        snap.infra_hosts + snap.apm_hosts + snap.profiled_hosts
        + snap.fargate_tasks + snap.profiled_fargate + snap.apm_fargate
    )
    cx.container_count = snap.containers + snap.profiled_containers

    cx.host_container_ts = (cx.host_count + cx.container_count) * TS_PER_HOST

    # Serverless metrics: daily functions/invocations × 3 labels × 10% unique dimensions
    sw_func_daily  = snap.serverless_functions / DAYS_PER_MONTH
    sw_invoc_daily = snap.serverless_invocations / DAYS_PER_MONTH
    cx.sw_func_ts  = sw_func_daily  * SW_LABEL_FACTOR
    cx.sw_invoc_ts = sw_invoc_daily * SW_LABEL_FACTOR

    cx.total_ts           = cx.host_container_ts + snap.custom_metrics + cx.sw_func_ts + cx.sw_invoc_ts
    cx.metrics_units_per_day = cx.total_ts * TS_TO_UNITS

    # ── Tracing ──────────────────────────────────────────────────────────────
    cx.ingested_spans_gb_month = snap.ingested_spans_bytes / 1e9

    # Indexed spans GB = (count + custom_events) × avg_span_size_kb / 1024² (KB→GB)
    cx.indexed_spans_gb_month = (
        (snap.indexed_spans + snap.custom_events) * AVG_SPAN_SIZE_KB
    ) / (1024 * 1024)

    if cx.ingested_spans_gb_month > 0:
        cx.indexed_pct_spans = cx.indexed_spans_gb_month / cx.ingested_spans_gb_month

    cx.daily_spans_ingest_gb  = cx.ingested_spans_gb_month / DAYS_PER_MONTH
    cx.daily_spans_indexed_gb = cx.daily_spans_ingest_gb * cx.indexed_pct_spans
    cx.daily_spans_archive_gb = cx.daily_spans_ingest_gb - cx.daily_spans_indexed_gb

    # ── RUM ──────────────────────────────────────────────────────────────────
    cx.rum_sessions_monthly = snap.rum_sessions
    cx.rum_errors_per_day   = snap.rum_errors / DAYS_PER_MONTH

    return cx


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 6 — Formatting helpers
# ═══════════════════════════════════════════════════════════════════════════

def _fmt(n: float | None, decimals: int = 1, suffix: str = "") -> str:
    if n is None:
        return "N/A"
    if n == 0:
        return f"0{suffix}"
    if abs(n) >= 1e12:
        return f"{n/1e12:.{decimals}f}T{suffix}"
    if abs(n) >= 1e9:
        return f"{n/1e9:.{decimals}f}B{suffix}"
    if abs(n) >= 1e6:
        return f"{n/1e6:.{decimals}f}M{suffix}"
    if abs(n) >= 1e3:
        return f"{n/1e3:.{decimals}f}K{suffix}"
    return f"{n:.{decimals}f}{suffix}"

def _bytes_to_tb(b: float) -> str:
    return f"{b/1e12:.2f} TB"

def _bytes_to_gb(b: float) -> str:
    return f"{b/1e9:.2f} GB"

def _usd(n: float | None) -> str:
    return "N/A" if n is None else f"${n:,.2f}"


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 7 — CSV output
# ═══════════════════════════════════════════════════════════════════════════

def generate_csv(snap: UsageSnapshot, cx: CoralogixSizing) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)

    w.writerow(["Datadog Usage Report", snap.month])
    w.writerow(["Account", snap.account_name or "N/A"])
    w.writerow(["Site", snap.site])
    w.writerow(["Pulled at", snap.pulled_at])
    w.writerow(["Note", FRESHNESS_NOTE])
    w.writerow([])

    # ── Bill Overview tiles ───────────────────────────────────────────────
    w.writerow(["=== DATADOG BILL OVERVIEW ==="])
    w.writerow(["Metric", "Value", "Unit"])

    tiles = [
        ("Infra Hosts",                    snap.infra_hosts,          "hosts (avg concurrent)"),
        ("APM Hosts",                      snap.apm_hosts,            "hosts (avg concurrent)"),
        ("Custom Metrics",                 snap.custom_metrics,       "time series"),
        ("Ingested Custom Metrics",        snap.ingested_custom_metrics, "time series"),
        ("Indexed Logs (3 Day)",           snap.indexed_logs_3day,    "events"),
        ("Indexed Logs (7 Day)",           snap.indexed_logs_7day,    "events"),
        ("Indexed Logs (15 Day)",          snap.indexed_logs_15day,   "events"),
        ("Indexed Logs (30 Day)",          snap.indexed_logs_30day,   "events"),
        ("Indexed Logs (45 Day)",          snap.indexed_logs_45day,   "events"),
        ("Indexed Logs (60 Day)",          snap.indexed_logs_60day,   "events"),
        ("Indexed Logs (90 Day)",          snap.indexed_logs_90day,   "events"),
        ("Indexed Logs (180 Day)",         snap.indexed_logs_180day,  "events"),
        ("Indexed Logs (360 Day)",         snap.indexed_logs_360day,  "events"),
        ("Indexed Logs (Live Search)",     snap.indexed_logs_live,    "events"),
        ("Indexed Logs (Rehydrated)",      snap.indexed_logs_rehydrated, "events"),
        ("Ingested Logs",                  snap.ingested_logs_bytes,  "bytes"),
        ("Container Hours",                snap.container_hours,      "container-hours/month"),
        ("Containers (avg concurrent)",    snap.containers,           "containers"),
        ("Ingested Spans",                 snap.ingested_spans_bytes, "bytes"),
        ("Indexed Spans",                  snap.indexed_spans,        "events"),
        ("Profiled Hosts",                 snap.profiled_hosts,       "hosts"),
        ("Profiled Container Hours",       snap.profiled_containers,  "container-hours"),
        ("Serverless Workload Functions",  snap.serverless_functions, "functions"),
        ("Serverless Invocations",         snap.serverless_invocations, "invocations"),
        ("Serverless App Instances",       snap.serverless_app_instances, "instances"),
        ("Fargate Tasks",                  snap.fargate_tasks,        "tasks"),
        ("APM Fargate Tasks",              snap.apm_fargate,          "tasks"),
        ("Profiled Fargate Tasks",         snap.profiled_fargate,     "tasks"),
        ("Network Hosts",                  snap.network_hosts,        "hosts"),
        ("DBM Hosts",                      snap.dbm_hosts,            "hosts"),
        ("Synthetics API Test Runs",       snap.synthetics_api,       "test runs"),
        ("Synthetics Browser Test Runs",   snap.synthetics_browser,   "test runs"),
        ("RUM Investigate (Sessions)",     snap.rum_sessions,         "sessions"),
        ("RUM Measure (Lite Sessions)",    snap.rum_lite_sessions,    "sessions"),
        ("Session Replay",                 snap.rum_replay,           "sessions"),
        ("Error Tracking Events",          snap.rum_errors,           "events"),
        ("Incident Management Seats",      snap.incident_management_seats, "seats"),
        ("Custom Events",                  snap.custom_events,        "events"),
        ("Product Analytics Sessions",     snap.product_analytics_sessions, "sessions"),
        ("Test Optimization Committers",   snap.test_optimization_committers, "committers"),
        ("Test Optimization Spans",        snap.test_optimization_spans, "spans"),
        ("App Builder Published Apps",     snap.app_builder_apps,     "apps"),
        ("Bits AI SRE Investigations",     snap.bits_ai_investigations, "investigations"),
        ("SIEM/Security Logs",             snap.security_logs_bytes,  "bytes"),
    ]
    for name, value, unit in tiles:
        w.writerow([name, value, unit])

    w.writerow([])

    # ── Coralogix sizing ─────────────────────────────────────────────────
    w.writerow(["=== CORALOGIX SIZING ==="])
    w.writerow(["Assumption: avg log size (KB)",   AVG_LOG_SIZE_KB])
    w.writerow(["Assumption: avg span size (KB)",  AVG_SPAN_SIZE_KB])
    w.writerow(["Assumption: TS per host/container", TS_PER_HOST])
    w.writerow(["Assumption: TS-to-Units factor",  TS_TO_UNITS])
    w.writerow(["Assumption: days per month",       DAYS_PER_MONTH])
    w.writerow([])

    w.writerow(["-- Logs --"])
    w.writerow(["Total Ingested Logs (GB/month)",   f"{cx.total_ingested_logs_gb_month:.2f}"])
    w.writerow(["Total Indexed Logs (events/month)", f"{cx.total_indexed_logs_count:.0f}"])
    w.writerow(["Total Indexed Logs Size (GB/month)", f"{cx.indexed_logs_size_gb_month:.2f}"])
    w.writerow(["Indexed Percentage",               f"{cx.indexed_pct_logs*100:.2f}%"])
    w.writerow(["Daily Ingested Logs (GB/day)",     f"{cx.daily_logs_gb:.2f}"])
    w.writerow(["  Monitoring 70% (GB/day)",        f"{cx.daily_logs_mon_gb:.2f}"])
    w.writerow(["  Compliance 30% (GB/day)",        f"{cx.daily_logs_comp_gb:.2f}"])
    w.writerow([])

    w.writerow(["-- Metrics --"])
    w.writerow(["Host count (all types)",           f"{cx.host_count:.0f}"])
    w.writerow(["Container count",                  f"{cx.container_count:.0f}"])
    w.writerow(["Host+Container TimeSeries",        f"{cx.host_container_ts:.0f}"])
    w.writerow(["Serverless Functions TS",          f"{cx.sw_func_ts:.2f}"])
    w.writerow(["Serverless Invocations TS",        f"{cx.sw_invoc_ts:.2f}"])
    w.writerow(["Total TimeSeries (NumSeries)",     f"{cx.total_ts:.0f}"])
    w.writerow(["Metrics Units/day",               f"{cx.metrics_units_per_day:.2f}"])
    w.writerow([])

    w.writerow(["-- Tracing --"])
    w.writerow(["Ingested Spans (GB/month)",        f"{cx.ingested_spans_gb_month:.2f}"])
    w.writerow(["Indexed Spans (GB/month)",         f"{cx.indexed_spans_gb_month:.2f}"])
    w.writerow(["Indexed Span Percentage",          f"{cx.indexed_pct_spans*100:.4f}%"])
    w.writerow(["Daily Ingested Spans (GB/day)",    f"{cx.daily_spans_ingest_gb:.2f}"])
    w.writerow(["  Indexed (GB/day)",               f"{cx.daily_spans_indexed_gb:.4f}"])
    w.writerow(["  Archive (GB/day)",               f"{cx.daily_spans_archive_gb:.2f}"])
    w.writerow([])

    w.writerow(["-- RUM --"])
    w.writerow(["RUM Sessions/month",               f"{cx.rum_sessions_monthly:.0f}"])
    w.writerow(["RUM Errors/day",                   f"{cx.rum_errors_per_day:.2f}"])
    w.writerow([])

    w.writerow(["-- Summary for TCO Calculator --"])
    w.writerow(["Logs GB/day",          f"{cx.daily_logs_gb:.2f}"])
    w.writerow(["Metrics NumSeries",    f"{cx.total_ts:.0f}"])
    w.writerow(["Tracing GB/day",       f"{cx.daily_spans_ingest_gb:.2f}"])
    w.writerow(["RUM Sessions/month",   f"{cx.rum_sessions_monthly:.0f}"])

    return buf.getvalue().encode()


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 8 — Excel output  (mirrors the template sheet layout)
# ═══════════════════════════════════════════════════════════════════════════

def generate_xlsx(snap: UsageSnapshot, cx: CoralogixSizing) -> bytes | None:
    if not HAS_OPENPYXL:
        return None

    wb = openpyxl.Workbook()

    # ── Helper styles ─────────────────────────────────────────────────────
    _PURPLE  = "632CA6"
    _ORANGE  = "FF5C35"
    _LGRAY   = "F5F5F5"
    _DGRAY   = "3C3C3C"
    _WHITE   = "FFFFFF"

    def hdr_cell(ws, row, col, value, bg=_PURPLE, fg=_WHITE, bold=True, sz=11):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(bold=bold, color=fg, size=sz)
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        return c

    def label_cell(ws, row, col, value, bold=False, bg=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(bold=bold, size=10)
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        return c

    def val_cell(ws, row, col, value, number_format=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(size=10)
        c.alignment = Alignment(horizontal="right")
        if number_format:
            c.number_format = number_format
        return c

    # ════════════════════════════════════════════════════════════════════
    # Sheet 1 — Bill Overview
    # ════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Bill Overview"
    ws1.column_dimensions["A"].width = 38
    ws1.column_dimensions["B"].width = 20
    ws1.column_dimensions["C"].width = 22
    ws1.column_dimensions["D"].width = 22
    ws1.row_dimensions[1].height = 28

    hdr_cell(ws1, 1, 1, f"Datadog Bill Overview — {snap.month}", bg=_PURPLE, sz=13)
    ws1.merge_cells("A1:D1")
    ws1.cell(row=2, column=1, value=f"Account: {snap.account_name or 'N/A'}  |  Site: {snap.site}  |  Pulled: {snap.pulled_at[:10]}")
    ws1.merge_cells("A2:D2")
    ws1.cell(row=3, column=1, value=f"Note: {FRESHNESS_NOTE}").font = Font(italic=True, size=9, color="666666")
    ws1.merge_cells("A3:D3")

    hdr_cell(ws1, 5, 1, "Product Metric",   bg=_DGRAY, sz=10)
    hdr_cell(ws1, 5, 2, "Raw Value",        bg=_DGRAY, sz=10)
    hdr_cell(ws1, 5, 3, "Formatted",        bg=_DGRAY, sz=10)
    hdr_cell(ws1, 5, 4, "Unit",             bg=_DGRAY, sz=10)

    tiles = [
        ("Infrastructure",           None,                          None,     ""),
        ("Infra Hosts",               snap.infra_hosts,              _fmt(snap.infra_hosts, 0), "hosts (avg concurrent)"),
        ("APM Hosts",                snap.apm_hosts,                _fmt(snap.apm_hosts, 0),   "hosts (avg concurrent)"),
        ("Container Hours",          snap.container_hours,          _fmt(snap.container_hours), "container-hours/month"),
        ("Containers (avg concurrent)", snap.containers,            _fmt(snap.containers, 0),  "containers"),
        ("Network Hosts",            snap.network_hosts,            _fmt(snap.network_hosts, 0), "hosts"),
        ("DBM Hosts",                snap.dbm_hosts,                _fmt(snap.dbm_hosts, 0),   "hosts"),
        ("",                         None,                          None,     ""),
        ("Custom Metrics",           None,                          None,     ""),
        ("Custom Metrics",           snap.custom_metrics,           _fmt(snap.custom_metrics), "time series"),
        ("Ingested Custom Metrics",  snap.ingested_custom_metrics,  _fmt(snap.ingested_custom_metrics), "time series"),
        ("",                         None,                          None,     ""),
        ("Logs",                     None,                          None,     ""),
        ("Ingested Logs",            snap.ingested_logs_bytes,      _bytes_to_tb(snap.ingested_logs_bytes), "bytes"),
        ("Indexed Logs (3 Day)",     snap.indexed_logs_3day,        _fmt(snap.indexed_logs_3day), "events"),
        ("Indexed Logs (7 Day)",     snap.indexed_logs_7day,        _fmt(snap.indexed_logs_7day), "events"),
        ("Indexed Logs (15 Day)",    snap.indexed_logs_15day,       _fmt(snap.indexed_logs_15day), "events"),
        ("Indexed Logs (30 Day)",    snap.indexed_logs_30day,       _fmt(snap.indexed_logs_30day), "events"),
        ("Indexed Logs (45 Day)",    snap.indexed_logs_45day,       _fmt(snap.indexed_logs_45day), "events"),
        ("Indexed Logs (60 Day)",    snap.indexed_logs_60day,       _fmt(snap.indexed_logs_60day), "events"),
        ("Indexed Logs (90 Day)",    snap.indexed_logs_90day,       _fmt(snap.indexed_logs_90day), "events"),
        ("Indexed Logs (180 Day)",   snap.indexed_logs_180day,      _fmt(snap.indexed_logs_180day), "events"),
        ("Indexed Logs (360 Day)",   snap.indexed_logs_360day,      _fmt(snap.indexed_logs_360day), "events"),
        ("Indexed Logs (Live Search)", snap.indexed_logs_live,      _fmt(snap.indexed_logs_live), "events"),
        ("Indexed Logs (Rehydrated)", snap.indexed_logs_rehydrated, _fmt(snap.indexed_logs_rehydrated), "events"),
        ("SIEM / Security Logs",     snap.security_logs_bytes,      _bytes_to_gb(snap.security_logs_bytes), "bytes"),
        ("",                         None,                          None,     ""),
        ("APM / Tracing",            None,                          None,     ""),
        ("Ingested Spans",           snap.ingested_spans_bytes,     _bytes_to_tb(snap.ingested_spans_bytes), "bytes"),
        ("Indexed Spans",            snap.indexed_spans,            _fmt(snap.indexed_spans), "events"),
        ("Profiled Hosts",           snap.profiled_hosts,           _fmt(snap.profiled_hosts, 0), "hosts"),
        ("Profiled Container Hours", snap.profiled_containers,      _fmt(snap.profiled_containers), "container-hours"),
        ("APM Fargate Tasks",        snap.apm_fargate,              _fmt(snap.apm_fargate, 0), "tasks"),
        ("Profiled Fargate Tasks",   snap.profiled_fargate,         _fmt(snap.profiled_fargate, 0), "tasks"),
        ("Custom Events",            snap.custom_events,            _fmt(snap.custom_events), "events"),
        ("",                         None,                          None,     ""),
        ("Serverless",               None,                          None,     ""),
        ("Serverless Workload Functions", snap.serverless_functions, _fmt(snap.serverless_functions, 0), "functions"),
        ("Serverless Invocations",   snap.serverless_invocations,   _fmt(snap.serverless_invocations), "invocations/month"),
        ("Serverless App Instances", snap.serverless_app_instances, _fmt(snap.serverless_app_instances, 0), "instances"),
        ("Fargate Tasks",            snap.fargate_tasks,            _fmt(snap.fargate_tasks, 0), "tasks"),
        ("",                         None,                          None,     ""),
        ("RUM",                      None,                          None,     ""),
        ("RUM Investigate",          snap.rum_sessions,             _fmt(snap.rum_sessions), "sessions/month"),
        ("RUM Measure",              snap.rum_lite_sessions,        _fmt(snap.rum_lite_sessions), "sessions/month"),
        ("Session Replay",           snap.rum_replay,               _fmt(snap.rum_replay), "sessions/month"),
        ("Error Tracking Events",    snap.rum_errors,               _fmt(snap.rum_errors), "events/month"),
        ("",                         None,                          None,     ""),
        ("Synthetics",               None,                          None,     ""),
        ("Synthetics API Test Runs",     snap.synthetics_api,       _fmt(snap.synthetics_api), "test runs/month"),
        ("Synthetics Browser Test Runs", snap.synthetics_browser,   _fmt(snap.synthetics_browser), "test runs/month"),
        ("",                         None,                          None,     ""),
        ("Other",                    None,                          None,     ""),
        ("Incident Management Seats",    snap.incident_management_seats,   _fmt(snap.incident_management_seats, 0), "seats"),
        ("Test Optimization Committers", snap.test_optimization_committers, _fmt(snap.test_optimization_committers, 0), "committers"),
        ("Test Optimization Spans",      snap.test_optimization_spans,     _fmt(snap.test_optimization_spans), "spans"),
        ("Product Analytics Sessions",   snap.product_analytics_sessions,  _fmt(snap.product_analytics_sessions), "sessions/month"),
        ("App Builder Published Apps",   snap.app_builder_apps,            _fmt(snap.app_builder_apps, 0), "apps"),
        ("Bits AI SRE Investigations",   snap.bits_ai_investigations,      _fmt(snap.bits_ai_investigations, 0), "investigations/month"),
    ]

    section_headers = {"Infrastructure", "Custom Metrics", "Logs", "APM / Tracing",
                       "Serverless", "RUM", "Synthetics", "Other"}

    for i, (name, raw, formatted, unit) in enumerate(tiles, start=6):
        row = i
        if name in section_headers:
            c = ws1.cell(row=row, column=1, value=name)
            c.font = Font(bold=True, size=10, color=_WHITE)
            c.fill = PatternFill("solid", fgColor=_ORANGE)
            ws1.merge_cells(f"A{row}:D{row}")
        elif name == "":
            pass
        else:
            fill = PatternFill("solid", fgColor=_LGRAY) if row % 2 == 0 else None
            c = ws1.cell(row=row, column=1, value=name)
            c.font = Font(size=10)
            if fill:
                c.fill = fill

            c2 = ws1.cell(row=row, column=2, value=raw if raw is not None else "")
            c2.font = Font(size=10)
            c2.alignment = Alignment(horizontal="right")
            if fill:
                c2.fill = fill

            c3 = ws1.cell(row=row, column=3, value=formatted or "")
            c3.font = Font(size=10, bold=True)
            c3.alignment = Alignment(horizontal="right")
            if fill:
                c3.fill = fill

            c4 = ws1.cell(row=row, column=4, value=unit)
            c4.font = Font(size=10, color="666666")
            if fill:
                c4.fill = fill

    # ════════════════════════════════════════════════════════════════════
    # Sheet 2 — Coralogix Sizing  (mirrors the Excel template layout)
    # ════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Coralogix Sizing")
    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 36
    ws2.row_dimensions[1].height = 28

    hdr_cell(ws2, 1, 1, f"Coralogix Sizing Estimate — {snap.month}", bg=_ORANGE, sz=13)
    ws2.merge_cells("A1:D1")
    ws2.cell(row=2, column=1, value=f"Based on Datadog usage for {snap.account_name or 'account'}. "
             "All assumptions are listed below and can be adjusted.")
    ws2.merge_cells("A2:D2")

    # ── Assumptions block ────────────────────────────────────────────────
    hdr_cell(ws2, 4, 1, "Sizing Assumptions", bg=_DGRAY, sz=10)
    ws2.merge_cells("A4:D4")
    assumptions = [
        ("Avg log line size", AVG_LOG_SIZE_KB, "KB"),
        ("Avg span size",     AVG_SPAN_SIZE_KB, "KB"),
        ("Time series per host/container", TS_PER_HOST, "TS"),
        ("TS-to-Units factor (metrics)",   TS_TO_UNITS, "Units/TS"),
        ("Days per month",    DAYS_PER_MONTH, "days"),
        ("Log tier split",    f"{int(LOG_TIER_MON*100)}% Mon / {int(LOG_TIER_COMP*100)}% Comp", ""),
    ]
    for i, (label, val, unit) in enumerate(assumptions, start=5):
        ws2.cell(row=i, column=1, value=label).font = Font(size=10)
        ws2.cell(row=i, column=2, value=val).font = Font(size=10, bold=True)
        ws2.cell(row=i, column=3, value=unit).font = Font(size=10, color="666666")

    # ── DD Usage Input block (mirrors template rows 1-10) ────────────────
    hdr_cell(ws2, 12, 1, "DD Usage Inputs", bg=_DGRAY, sz=10)
    ws2.merge_cells("A12:D12")

    dd_inputs = [
        ("Infra Hosts",                         snap.infra_hosts,            "hosts (avg concurrent)"),
        ("APM Hosts",                           snap.apm_hosts,              "hosts (avg concurrent)"),
        ("Profiled Hosts",                      snap.profiled_hosts,         "hosts (avg concurrent)"),
        ("Containers",                          snap.containers,             "containers (avg concurrent)"),
        ("Custom Metrics",                      snap.custom_metrics,         "time series"),
        ("Ingested Logs",                       snap.ingested_logs_bytes/1e9,"GB/month"),
        ("Security Logs",                       snap.security_logs_bytes/1e9,"GB/month"),
        ("Indexed Logs — Live & Rehydrated",    (snap.indexed_logs_3day+snap.indexed_logs_7day+snap.indexed_logs_live+snap.indexed_logs_rehydrated)/1e6, "Million events"),
        ("Indexed Logs — 15 Day",               snap.indexed_logs_15day/1e6, "Million events"),
        ("Indexed Logs — 90 Day (30-360d sum)", (snap.indexed_logs_30day+snap.indexed_logs_45day+snap.indexed_logs_60day+snap.indexed_logs_90day+snap.indexed_logs_180day+snap.indexed_logs_360day)/1e6, "Million events"),
        ("Ingested Spans",                      snap.ingested_spans_bytes/1e9,"GB/month"),
        ("Indexed Spans",                       snap.indexed_spans/1e6,       "Million events"),
        ("Custom Events",                       snap.custom_events,           "events/month"),
        ("Serverless Functions (monthly avg)",  snap.serverless_functions,    "functions"),
        ("Serverless Invocations (monthly)",    snap.serverless_invocations,  "invocations"),
        ("Fargate Tasks",                       snap.fargate_tasks,           "tasks"),
        ("APM Fargate Tasks",                   snap.apm_fargate,             "tasks"),
        ("Profiled Fargate Tasks",              snap.profiled_fargate,        "tasks"),
        ("RUM Sessions",                        snap.rum_sessions,            "sessions/month"),
        ("Error Tracking Events",               snap.rum_errors,              "events/month"),
    ]
    hdr_cell(ws2, 13, 1, "Input Metric", bg=_DGRAY, sz=10)
    hdr_cell(ws2, 13, 2, "Value",        bg=_DGRAY, sz=10)
    hdr_cell(ws2, 13, 3, "Unit",         bg=_DGRAY, sz=10)
    for i, (label, val, unit) in enumerate(dd_inputs, start=14):
        fill = PatternFill("solid", fgColor=_LGRAY) if i % 2 == 0 else None
        c1 = ws2.cell(row=i, column=1, value=label)
        c1.font = Font(size=10)
        if fill: c1.fill = fill
        c2 = ws2.cell(row=i, column=2, value=round(val, 4) if isinstance(val, float) else val)
        c2.font = Font(size=10, bold=True)
        c2.alignment = Alignment(horizontal="right")
        if fill: c2.fill = fill
        c3 = ws2.cell(row=i, column=3, value=unit)
        c3.font = Font(size=10, color="666666")
        if fill: c3.fill = fill

    # ── Coralogix Sizing Output ───────────────────────────────────────────
    out_row = 14 + len(dd_inputs) + 2

    sections = [
        ("LOGS", _ORANGE, [
            ("Total Ingested Logs",           cx.total_ingested_logs_gb_month, "GB/month",   "ingested_logs_gb + security_logs_gb"),
            ("Total Indexed Logs (events)",   cx.total_indexed_logs_count,     "events/month","all retention tiers summed"),
            ("Total Indexed Logs Size",       cx.indexed_logs_size_gb_month,   "GB/month",   "count × avg_log_size_kb × 1024 ÷ 1024³"),
            ("Indexed Percentage",            cx.indexed_pct_logs * 100,       "%",          "indexed_size_gb / ingested_gb"),
            ("Daily Ingested Logs",           cx.daily_logs_gb,                "GB/day",     "ingested_gb ÷ 30"),
            ("  Monitoring (70%)",            cx.daily_logs_mon_gb,            "GB/day",     "daily × 0.70"),
            ("  Compliance (30%)",            cx.daily_logs_comp_gb,           "GB/day",     "daily × 0.30"),
        ]),
        ("METRICS", _ORANGE, [
            ("Host Count (all types incl. fargate)", cx.host_count,         "hosts",       "infra+apm+profiled+fargate types"),
            ("Container Count",               cx.container_count,              "containers", "containers + profiled_containers"),
            ("Host+Container TS",             cx.host_container_ts,            "NumSeries",  "(hosts+containers) × 750"),
            ("Serverless Functions TS",       cx.sw_func_ts,                   "NumSeries",  "daily_functions × 0.30"),
            ("Serverless Invocations TS",     cx.sw_invoc_ts,                  "NumSeries",  "daily_invocations × 0.30"),
            ("Total TimeSeries (NumSeries)",  cx.total_ts,                     "NumSeries",  "host_cont_ts + custom_metrics + sw_ts"),
            ("Metrics Units / Day",           cx.metrics_units_per_day,        "Units/day",  "total_ts × 3.3e-5"),
        ]),
        ("TRACING", _ORANGE, [
            ("Ingested Spans",                cx.ingested_spans_gb_month,      "GB/month",   "twol_ingested_events_bytes_sum ÷ 1e9"),
            ("Indexed Spans",                 cx.indexed_spans_gb_month,       "GB/month",   "(indexed_spans+custom_events) × 1.5 KB ÷ 1024²"),
            ("Indexed Span Percentage",       cx.indexed_pct_spans * 100,      "%",          "indexed_gb / ingested_gb"),
            ("Daily Ingested Spans",          cx.daily_spans_ingest_gb,        "GB/day",     "ingested_gb ÷ 30"),
            ("  Indexed (GB/day)",            cx.daily_spans_indexed_gb,       "GB/day",     "daily × indexed_pct"),
            ("  Archive (GB/day)",            cx.daily_spans_archive_gb,       "GB/day",     "daily - indexed_daily"),
        ]),
        ("RUM", _ORANGE, [
            ("RUM Sessions / Month",          cx.rum_sessions_monthly,         "sessions",   "rum_total_session_count"),
            ("RUM Errors / Day",              cx.rum_errors_per_day,           "events/day", "error_tracking_events ÷ 30"),
        ]),
        ("SUMMARY — Apply to Coralogix TCO Calculator", _PURPLE, [
            ("Logs GB/day",                   cx.daily_logs_gb,                "GB/day",     ""),
            ("Metrics NumSeries",             cx.total_ts,                     "NumSeries",  ""),
            ("Tracing GB/day",                cx.daily_spans_ingest_gb,        "GB/day",     ""),
            ("RUM Sessions/month",            cx.rum_sessions_monthly,         "sessions",   ""),
        ]),
    ]

    for section_name, section_color, rows in sections:
        hdr_cell(ws2, out_row, 1, section_name, bg=section_color, sz=10)
        ws2.merge_cells(f"A{out_row}:D{out_row}")
        out_row += 1
        hdr_cell(ws2, out_row, 1, "Metric",     bg=_DGRAY, sz=9)
        hdr_cell(ws2, out_row, 2, "Value",      bg=_DGRAY, sz=9)
        hdr_cell(ws2, out_row, 3, "Unit",       bg=_DGRAY, sz=9)
        hdr_cell(ws2, out_row, 4, "Formula",    bg=_DGRAY, sz=9)
        out_row += 1
        for i, (label, val, unit, formula) in enumerate(rows):
            fill = PatternFill("solid", fgColor=_LGRAY) if i % 2 == 0 else None
            c1 = ws2.cell(row=out_row, column=1, value=label)
            c1.font = Font(size=10)
            if fill: c1.fill = fill
            c2 = ws2.cell(row=out_row, column=2,
                          value=round(val, 4) if isinstance(val, float) else val)
            c2.font = Font(size=10, bold=True)
            c2.alignment = Alignment(horizontal="right")
            if fill: c2.fill = fill
            c3 = ws2.cell(row=out_row, column=3, value=unit)
            c3.font = Font(size=10, color="666666")
            if fill: c3.fill = fill
            c4 = ws2.cell(row=out_row, column=4, value=formula)
            c4.font = Font(size=9, color="888888", italic=True)
            if fill: c4.fill = fill
            out_row += 1
        out_row += 1

    # ════════════════════════════════════════════════════════════════════
    # Sheet 3 — Billable Breakdown
    # ════════════════════════════════════════════════════════════════════
    if snap.billable_breakdown:
        ws3 = wb.create_sheet("Billable Breakdown")
        ws3.column_dimensions["A"].width = 35
        ws3.column_dimensions["B"].width = 20
        ws3.column_dimensions["C"].width = 20
        ws3.column_dimensions["D"].width = 20
        ws3.column_dimensions["E"].width = 20
        ws3.column_dimensions["F"].width = 16

        hdr_cell(ws3, 1, 1, "Billable Usage Breakdown by Dimension", bg=_PURPLE, sz=12)
        ws3.merge_cells("A1:F1")
        hdr_cell(ws3, 2, 1, "Billing Dimension", bg=_DGRAY, sz=10)
        hdr_cell(ws3, 2, 2, "Billable Usage",    bg=_DGRAY, sz=10)
        hdr_cell(ws3, 2, 3, "Committed Usage",   bg=_DGRAY, sz=10)
        hdr_cell(ws3, 2, 4, "On-Demand Usage",   bg=_DGRAY, sz=10)
        hdr_cell(ws3, 2, 5, "Start Date",        bg=_DGRAY, sz=10)
        hdr_cell(ws3, 2, 6, "Unit",              bg=_DGRAY, sz=10)

        for i, row_data in enumerate(snap.billable_breakdown, start=3):
            fill = PatternFill("solid", fgColor=_LGRAY) if i % 2 == 0 else None
            for col, key in enumerate(
                ["billing_dimension", "billable", "committed", "on_demand", "start_date", "unit"], start=1
            ):
                c = ws3.cell(row=i, column=col, value=row_data.get(key, ""))
                c.font = Font(size=10)
                if col > 1:
                    c.alignment = Alignment(horizontal="right")
                if fill:
                    c.fill = fill

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 9 — HTML report  (self-contained, no external dependencies)
# ═══════════════════════════════════════════════════════════════════════════

def generate_html(snap: UsageSnapshot, cx: CoralogixSizing) -> bytes:

    def tile(label: str, value: str, sub: str = "", color: str = "#632CA6") -> str:
        return f"""
        <div class="tile">
          <div class="tile-label">{label}</div>
          <div class="tile-value" style="color:{color}">{value}</div>
          {f'<div class="tile-sub">{sub}</div>' if sub else ""}
        </div>"""

    def row(label: str, value: str, note: str = "", even: bool = False) -> str:
        bg = "#f9f9f9" if even else "#ffffff"
        return (f'<tr style="background:{bg}">'
                f'<td>{label}</td><td class="num">{value}</td>'
                f'<td class="note">{note}</td></tr>')

    # ── Bill Overview tiles ───────────────────────────────────────────────
    infra_tiles = "".join([
        tile("Infra Hosts",              _fmt(snap.infra_hosts, 0),    "avg concurrent"),
        tile("APM Hosts",                _fmt(snap.apm_hosts, 0),      "avg concurrent"),
        tile("Container Hours",          _fmt(snap.container_hours),   "host-hours/month"),
        tile("Containers (avg)",         _fmt(snap.containers, 0),     "avg concurrent"),
        tile("Network Hosts",            _fmt(snap.network_hosts, 0)),
        tile("DBM Hosts",                _fmt(snap.dbm_hosts, 0)),
        tile("Profiled Hosts",           _fmt(snap.profiled_hosts, 0)),
        tile("Profiled Containers (avg)",_fmt(snap.profiled_containers, 0), "avg concurrent"),
        tile("Fargate Tasks",            _fmt(snap.fargate_tasks, 0)),
        tile("APM Fargate Tasks",        _fmt(snap.apm_fargate, 0)),
    ])
    metrics_tiles = "".join([
        tile("Custom Metrics",           _fmt(snap.custom_metrics),           color="#1a56db"),
        tile("Ingested Custom Metrics",  _fmt(snap.ingested_custom_metrics),  color="#1a56db"),
    ])
    logs_tiles = "".join([
        tile("Ingested Logs",            _bytes_to_tb(snap.ingested_logs_bytes)),
        tile("Indexed Logs (3 Day)",     _fmt(snap.indexed_logs_3day)),
        tile("Indexed Logs (7 Day)",     _fmt(snap.indexed_logs_7day)),
        tile("Indexed Logs (15 Day)",    _fmt(snap.indexed_logs_15day)),
        tile("Indexed Logs (30 Day)",    _fmt(snap.indexed_logs_30day)),
        tile("Indexed Logs (45 Day)",    _fmt(snap.indexed_logs_45day)),
        tile("Indexed Logs (60 Day)",    _fmt(snap.indexed_logs_60day)),
        tile("Indexed Logs (90 Day)",    _fmt(snap.indexed_logs_90day)),
        tile("Indexed Logs (180 Day)",   _fmt(snap.indexed_logs_180day)),
        tile("Indexed Logs (360 Day)",   _fmt(snap.indexed_logs_360day)),
        tile("Indexed Logs (Live)",      _fmt(snap.indexed_logs_live)),
        tile("Indexed Logs (Rehydrated)",_fmt(snap.indexed_logs_rehydrated)),
        tile("SIEM / Security Logs",     _bytes_to_gb(snap.security_logs_bytes)),
    ])
    apm_tiles = "".join([
        tile("Ingested Spans",           _bytes_to_tb(snap.ingested_spans_bytes)),
        tile("Indexed Spans",            _fmt(snap.indexed_spans)),
        tile("Custom Events",            _fmt(snap.custom_events)),
    ])
    sl_tiles = "".join([
        tile("Serverless Functions",     _fmt(snap.serverless_functions, 0)),
        tile("Serverless Invocations",   _fmt(snap.serverless_invocations)),
        tile("Serverless App Instances", _fmt(snap.serverless_app_instances, 0)),
    ])
    rum_tiles = "".join([
        tile("RUM Investigate",          _fmt(snap.rum_sessions)),
        tile("RUM Measure",              _fmt(snap.rum_lite_sessions)),
        tile("Session Replay",           _fmt(snap.rum_replay)),
        tile("Error Tracking Events",    _fmt(snap.rum_errors)),
    ])
    synth_tiles = "".join([
        tile("Synthetics API",           _fmt(snap.synthetics_api)),
        tile("Synthetics Browser",       _fmt(snap.synthetics_browser)),
    ])
    other_tiles = "".join([
        tile("Incident Mgmt Seats",      _fmt(snap.incident_management_seats, 0)),
        tile("Test Opt. Committers",     _fmt(snap.test_optimization_committers, 0)),
        tile("Test Opt. Spans",          _fmt(snap.test_optimization_spans)),
        tile("Product Analytics",        _fmt(snap.product_analytics_sessions)),
        tile("App Builder Apps",         _fmt(snap.app_builder_apps, 0)),
        tile("Bits AI Investigations",   _fmt(snap.bits_ai_investigations, 0)),
    ])

    # ── Coralogix sizing rows ─────────────────────────────────────────────
    cx_rows_logs = "".join([
        row("Total Ingested Logs",           f"{cx.total_ingested_logs_gb_month:.2f} GB/month",  "", True),
        row("Total Indexed Logs",            f"{cx.total_indexed_logs_count:,.0f} events",        "all retention tiers"),
        row("Indexed Logs Size",             f"{cx.indexed_logs_size_gb_month:.2f} GB/month",    f"avg {AVG_LOG_SIZE_KB} KB/log", True),
        row("Indexed Percentage",            f"{cx.indexed_pct_logs*100:.2f}%",                  "indexed size / ingested"),
        row("Daily Ingested Logs",           f"{cx.daily_logs_gb:.2f} GB/day",                   "ingested ÷ 30", True),
        row("  → Monitoring (70%)",          f"{cx.daily_logs_mon_gb:.2f} GB/day",               "", True),
        row("  → Compliance (30%)",          f"{cx.daily_logs_comp_gb:.2f} GB/day",              ""),
    ])
    cx_rows_metrics = "".join([
        row("Host Count (all types)",        f"{cx.host_count:,.0f}",                            "infra+apm+profiled+fargate", True),
        row("Container Count",               f"{cx.container_count:,.0f}",                       "containers+profiled"),
        row("Host+Container TimeSeries",     f"{cx.host_container_ts:,.0f}",                     f"× {TS_PER_HOST} TS each", True),
        row("Serverless Functions TS",       f"{cx.sw_func_ts:,.2f}",                            "daily_funcs × 0.30"),
        row("Serverless Invocations TS",     f"{cx.sw_invoc_ts:,.2f}",                           "daily_invoc × 0.30", True),
        row("Total TimeSeries (NumSeries)",  f"{cx.total_ts:,.0f}",                             "sum of all TS"),
        row("Metrics Units / Day",           f"{cx.metrics_units_per_day:.2f}",                  f"total_ts × {TS_TO_UNITS}", True),
    ])
    cx_rows_tracing = "".join([
        row("Ingested Spans",                f"{cx.ingested_spans_gb_month:.2f} GB/month",       "", True),
        row("Indexed Spans",                 f"{cx.indexed_spans_gb_month:.4f} GB/month",        f"avg {AVG_SPAN_SIZE_KB} KB/span"),
        row("Indexed Span Percentage",       f"{cx.indexed_pct_spans*100:.4f}%",                 "", True),
        row("Daily Ingested Spans",          f"{cx.daily_spans_ingest_gb:.2f} GB/day",           "ingested ÷ 30"),
        row("  → Indexed",                   f"{cx.daily_spans_indexed_gb:.4f} GB/day",          "", True),
        row("  → Archive",                   f"{cx.daily_spans_archive_gb:.2f} GB/day",          ""),
    ])
    cx_rows_rum = "".join([
        row("RUM Sessions / Month",          f"{cx.rum_sessions_monthly:,.0f}",                  "", True),
        row("RUM Errors / Day",              f"{cx.rum_errors_per_day:,.2f}",                    "error_tracking ÷ 30"),
    ])

    cost_section = ""
    cost_section = ""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Datadog → Coralogix Report — {snap.month}</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
          background: #f0f0f0; color: #1a1a1a; font-size: 14px; }}
  header {{ background: #632CA6; color: #fff; padding: 24px 32px; }}
  header h1 {{ font-size: 22px; font-weight: 700; margin-bottom: 4px; }}
  header p  {{ font-size: 13px; opacity: 0.85; }}
  .freshness {{ background: #fff8e1; border-left: 4px solid #f59e0b;
               padding: 10px 16px; margin: 16px 32px; font-size: 12px; color: #92400e; }}
  main {{ padding: 16px 32px 40px; }}
  section {{ margin-bottom: 32px; }}
  h2 {{ font-size: 15px; font-weight: 700; margin-bottom: 12px;
        padding: 8px 12px; background: #FF5C35; color: #fff; border-radius: 4px; }}
  .tile-grid {{ display: flex; flex-wrap: wrap; gap: 10px; }}
  .tile {{ background: #fff; border: 1px solid #e0e0e0; border-radius: 6px;
           padding: 12px 16px; min-width: 160px; flex: 1 1 160px; max-width: 220px; }}
  .tile-label {{ font-size: 11px; color: #666; margin-bottom: 4px; }}
  .tile-value {{ font-size: 20px; font-weight: 700; color: #632CA6; }}
  .tile-sub   {{ font-size: 11px; color: #999; margin-top: 2px; }}
  table {{ width: 100%; border-collapse: collapse; background: #fff;
           border: 1px solid #e0e0e0; border-radius: 6px; overflow: hidden; }}
  th {{ background: #3c3c3c; color: #fff; text-align: left; padding: 8px 12px; font-size: 12px; }}
  td {{ padding: 7px 12px; border-bottom: 1px solid #f0f0f0; font-size: 13px; }}
  td.num  {{ text-align: right; font-weight: 600; font-variant-numeric: tabular-nums; }}
  td.note {{ font-size: 11px; color: #888; }}
  .summary-box {{ background: #fff; border: 2px solid #FF5C35; border-radius: 8px;
                 padding: 20px 24px; display: flex; flex-wrap: wrap; gap: 24px; }}
  .summary-item {{ flex: 1 1 180px; }}
  .summary-item .s-label {{ font-size: 11px; color: #888; text-transform: uppercase;
                             letter-spacing: 0.05em; margin-bottom: 4px; }}
  .summary-item .s-value {{ font-size: 24px; font-weight: 800; color: #FF5C35; }}
  .summary-item .s-unit  {{ font-size: 12px; color: #666; margin-top: 2px; }}
  footer {{ text-align: center; padding: 16px; font-size: 11px; color: #999; }}
</style>
</head>
<body>
<header>
  <h1>Datadog → Coralogix Usage Report</h1>
  <p>Month: <strong>{snap.month}</strong> &nbsp;|&nbsp;
     Account: <strong>{snap.account_name or "N/A"}</strong> &nbsp;|&nbsp;
     Site: <strong>{snap.site}</strong> &nbsp;|&nbsp;
     Generated: <strong>{snap.pulled_at[:19]} UTC</strong></p>
</header>

<div class="freshness">
  <strong>Data freshness note:</strong> {FRESHNESS_NOTE}
</div>

<main>

<!-- ── TCO Summary ─────────────────────────────────────────────────────── -->
<section>
  <h2>Coralogix TCO Summary</h2>
  <p style="font-size:12px;color:#666;margin-bottom:12px;">
    Apply these four numbers to the Coralogix TCO Calculator.
  </p>
  <div class="summary-box">
    <div class="summary-item">
      <div class="s-label">Logs</div>
      <div class="s-value">{cx.daily_logs_gb:.2f}</div>
      <div class="s-unit">GB / day</div>
    </div>
    <div class="summary-item">
      <div class="s-label">Metrics</div>
      <div class="s-value">{_fmt(cx.total_ts, 1)}</div>
      <div class="s-unit">NumSeries (TimeSeries)</div>
    </div>
    <div class="summary-item">
      <div class="s-label">Tracing</div>
      <div class="s-value">{cx.daily_spans_ingest_gb:.2f}</div>
      <div class="s-unit">GB / day</div>
    </div>
    <div class="summary-item">
      <div class="s-label">RUM Sessions</div>
      <div class="s-value">{_fmt(cx.rum_sessions_monthly, 1)}</div>
      <div class="s-unit">sessions / month</div>
    </div>
  </div>
</section>

<!-- ── Bill Overview ──────────────────────────────────────────────────── -->
<section>
  <h2>Infrastructure</h2>
  <div class="tile-grid">{infra_tiles}</div>
</section>
<section>
  <h2>Custom Metrics</h2>
  <div class="tile-grid">{metrics_tiles}</div>
</section>
<section>
  <h2>Logs</h2>
  <div class="tile-grid">{logs_tiles}</div>
</section>
<section>
  <h2>APM / Tracing</h2>
  <div class="tile-grid">{apm_tiles}</div>
</section>
<section>
  <h2>Serverless</h2>
  <div class="tile-grid">{sl_tiles}</div>
</section>
<section>
  <h2>RUM</h2>
  <div class="tile-grid">{rum_tiles}</div>
</section>
<section>
  <h2>Synthetics</h2>
  <div class="tile-grid">{synth_tiles}</div>
</section>
<section>
  <h2>Other Products</h2>
  <div class="tile-grid">{other_tiles}</div>
</section>

{cost_section}

<!-- ── Coralogix Sizing Detail ────────────────────────────────────────── -->
<section>
  <h2>Coralogix Sizing — Logs</h2>
  <table>
    <thead><tr><th>Metric</th><th class="num">Value</th><th class="note">Notes</th></tr></thead>
    <tbody>{cx_rows_logs}</tbody>
  </table>
</section>
<section>
  <h2>Coralogix Sizing — Metrics</h2>
  <table>
    <thead><tr><th>Metric</th><th class="num">Value</th><th class="note">Notes</th></tr></thead>
    <tbody>{cx_rows_metrics}</tbody>
  </table>
</section>
<section>
  <h2>Coralogix Sizing — Tracing</h2>
  <table>
    <thead><tr><th>Metric</th><th class="num">Value</th><th class="note">Notes</th></tr></thead>
    <tbody>{cx_rows_tracing}</tbody>
  </table>
</section>
<section>
  <h2>Coralogix Sizing — RUM</h2>
  <table>
    <thead><tr><th>Metric</th><th class="num">Value</th><th class="note">Notes</th></tr></thead>
    <tbody>{cx_rows_rum}</tbody>
  </table>
</section>

<!-- ── Sizing Assumptions ────────────────────────────────────────────── -->
<section>
  <h2>Sizing Assumptions</h2>
  <table>
    <thead><tr><th>Parameter</th><th class="num">Value</th></tr></thead>
    <tbody>
      {row("Average log line size",         f"{AVG_LOG_SIZE_KB} KB",       "", True)}
      {row("Average span size",             f"{AVG_SPAN_SIZE_KB} KB")}
      {row("Time series per host/container",f"{TS_PER_HOST}",              "TS", True)}
      {row("TS-to-Units factor",            f"{TS_TO_UNITS}",              "Coralogix metrics conversion")}
      {row("Days per month",                f"{int(DAYS_PER_MONTH)}",      "", True)}
      {row("Log tier split",f"{int(LOG_TIER_MON*100)}% Mon / {int(LOG_TIER_COMP*100)}% Comp")}
    </tbody>
  </table>
</section>

</main>
<footer>
  Datadog → Coralogix Usage Report · {snap.month} ·
  Generated {snap.pulled_at[:10]} · dd_usage_pull.py
</footer>
</body>
</html>"""

    return html.encode("utf-8")


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 10 — ZIP packaging
# ═══════════════════════════════════════════════════════════════════════════

def build_zip(
    month: str,
    raw_json: bytes,
    csv_data: bytes,
    xlsx_data: bytes | None,
    html_data: bytes,
    out_dir: Path,
) -> Path:
    zip_name = f"datadog_coralogix_report_{month}.zip"
    zip_path  = out_dir / zip_name

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"datadog_raw_{month}.json",         raw_json)
        zf.writestr(f"datadog_usage_{month}.csv",        csv_data)
        zf.writestr(f"report_{month}.html",              html_data)
        if xlsx_data:
            zf.writestr(f"coralogix_sizing_{month}.xlsx", xlsx_data)

    return zip_path


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 11 — Main
# ═══════════════════════════════════════════════════════════════════════════

def _previous_month() -> str:
    today = datetime.now(timezone.utc)
    first = today.replace(day=1)
    last_month = first - timedelta(days=1)
    return last_month.strftime("%Y-%m")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Pull Datadog usage and convert to Coralogix sizing.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--month", default=os.getenv("DD_MONTH", _previous_month()),
                        help="Target month YYYY-MM (default: previous month)")
    parser.add_argument("--site",  default=os.getenv("DD_SITE", "datadoghq.com"),
                        help="Datadog site (default: datadoghq.com)")
    parser.add_argument("--out",   default=".", help="Output directory (default: current dir)")
    args = parser.parse_args()

    month   = args.month
    site    = args.site
    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    api_key = os.getenv("DD_API_KEY", "").strip()
    app_key = os.getenv("DD_APP_KEY", os.getenv("DD_APPLICATION_KEY", "")).strip()

    if not api_key:
        sys.exit("\n  DD_API_KEY is not set. Add it to your .env file or environment.\n")
    if not app_key:
        sys.exit("\n  DD_APP_KEY is not set. Add it to your .env file or environment.\n")

    print(f"\n  Datadog → Coralogix Usage Report")
    print(f"  Site  : {site}")
    print(f"  Month : {month}")
    print(f"  Output: {out_dir}\n")

    client = DatadogClient(api_key, app_key, site)

    # ── Fetch all data ────────────────────────────────────────────────────
    summary_raw        = None
    billable_raw       = None
    print("  [1/2] Fetching usage summary …")
    try:
        summary_raw = client.usage_summary(month, month)
        print("        OK")
    except Exception as exc:
        print(f"        WARN: {exc}")
        summary_raw = {}

    print("  [2/2] Fetching billable usage summary …")
    try:
        billable_raw = client.billable_summary(month)
        print("        OK")
    except PermissionError as exc:
        print(f"        SKIP (billing_read not available): {exc}")
    except Exception as exc:
        print(f"        WARN: {exc}")

    # ── Process ───────────────────────────────────────────────────────────
    print("\n  Processing …")
    snap = extract_usage_snapshot(
        summary_raw or {},
        billable_raw,
        month,
        site,
    )
    cx = compute_coralogix_sizing(snap)

    # ── Generate outputs ──────────────────────────────────────────────────
    print("  Generating outputs …")
    raw_json  = json.dumps(snap.raw, indent=2, default=str).encode()
    csv_data  = generate_csv(snap, cx)
    xlsx_data = generate_xlsx(snap, cx)
    html_data = generate_html(snap, cx)

    zip_path = build_zip(month, raw_json, csv_data, xlsx_data, html_data, out_dir)

    # ── Print summary ─────────────────────────────────────────────────────
    print(f"\n  ══════════════════════════════════════════")
    print(f"  Coralogix TCO Sizing Summary — {month}")
    print(f"  ══════════════════════════════════════════")
    print(f"  Logs     : {cx.daily_logs_gb:.2f} GB/day")
    print(f"  Metrics  : {cx.total_ts:,.0f} NumSeries")
    print(f"  Tracing  : {cx.daily_spans_ingest_gb:.2f} GB/day")
    print(f"  RUM      : {cx.rum_sessions_monthly:,.0f} sessions/month")
    print(f"  ══════════════════════════════════════════")
    print(f"\n  Output ZIP : {zip_path}")
    print(f"  Contents   :")
    print(f"    datadog_raw_{month}.json")
    print(f"    datadog_usage_{month}.csv")
    if xlsx_data:
        print(f"    coralogix_sizing_{month}.xlsx")
    print(f"    report_{month}.html")
    print()


if __name__ == "__main__":
    main()
