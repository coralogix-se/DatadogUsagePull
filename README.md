# Datadog → Coralogix Usage Report

Pulls Datadog usage from the official Usage Metering API, converts it into
Coralogix TCO sizing inputs (and Checkly inputs for Synthetics), and packages
everything into a ZIP for handoff.

The script always pulls a **3-month window**: the target month plus the two
months before it. That powers month-over-month trend analysis in the HTML and
Excel outputs.

---

## What it produces

One ZIP named like:

```text
datadog_coralogix_report_2026-01_to_2026-03.zip
```

| File | Contents |
|---|---|
| `datadog_raw_<range>.json` | Combined raw usage responses for all months in the window |
| `datadog_usage_YYYY-MM.csv` | One CSV per month (Bill Overview + sizing + TCO/Checkly summaries) |
| `coralogix_sizing_<range>.xlsx` | Excel workbook: Bill Overview, Coralogix sizing, Trend Analysis sheet |
| `report_<range>.html` | Self-contained HTML report (open in any browser) |

### HTML report — what to enter where

**Coralogix TCO sheet** (latest month in the window):

| Input | Unit | Notes |
|---|---|---|
| Logs | GB / day | Split **70% Monitoring / 30% Compliance** |
| Metrics | NumSeries | Hosts + containers + custom metrics (+ serverless TS) |
| Tracing | GB / day | Split **10% Monitoring / 90% Compliance** |
| RUM Sessions | sessions / day | Total sessions ÷ 30 |
| RUM Recording | recordings / day | Session replay ÷ 30 |

**Checkly** (from Datadog Synthetics):

| Input | Unit |
|---|---|
| API checks | test runs / day |
| Browser checks | test runs / day |

The report also includes:

- Growth trend table across the 3-month window (values + MoM %)
- Datadog source tiles (infra, logs, tracing, RUM, synthetics)
- Sizing detail tables and the conversion assumptions used

---

## Requirements

- Python 3.9+
- Datadog **API key** with `usage_read`
- Datadog **Application key** with `usage_read`

No `billing_read` or cost endpoints are used. The script collects **usage volumes only** — not how much you pay Datadog.

---

## Quick start

### 1. Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Configure

```bash
cp .env.example .env
```

```env
DD_API_KEY=<your API key>
DD_APP_KEY=<your Application key>
DD_SITE=datadoghq.com        # or datadoghq.eu, us3.datadoghq.com, etc.
DD_MONTH=2026-03             # target month; script also pulls the 2 months before
```

### 3. Run

```bash
python dd_usage_pull.py
```

Example output:

```text
  Datadog → Coralogix Usage Report
  Site    : datadoghq.com
  Range   : 2026-01 → 2026-03  (3 months)

  [1/1] Fetching usage summary (2026-04 → 2026-06) …

  Coralogix TCO Sizing — Latest Month: 2026-03
  Logs     : 11092.61 GB/day
  Metrics  : 36,307,824 NumSeries
  Tracing  : 625.90 GB/day
  RUM      : 61,229 sessions/day  |  61,020 recordings/day
  Checkly  : 457,743 API/day  |  8,709 browser/day

  Output ZIP : ./datadog_coralogix_report_2026-01_to_2026-03.zip
```

Send the ZIP to your Coralogix contact.

---

## CLI options

```bash
python dd_usage_pull.py --month 2026-03 --site datadoghq.eu --out /tmp/reports
```

| Option | Default | Description |
|---|---|---|
| `--month` | previous calendar month | Target month (`YYYY-MM`). Also pulls the 2 months before it |
| `--site` | `datadoghq.com` | Datadog site |
| `--out` | `.` | Output directory for the ZIP |

Precedence: CLI flags → environment variables → `.env` file.

---

## Datadog API endpoints used

Official [Usage Metering API](https://docs.datadoghq.com/api/latest/usage-metering/) only — no UI scraping, no cost endpoints.

| Endpoint | Purpose |
|---|---|
| `GET /api/v1/usage/summary` | Product usage volumes for the 3-month window |

The script does **not** call `estimated_cost`, `historical_cost`, `projected_cost`, or `billable-summary`.

> Datadog usage data can lag up to ~72 hours.

### Important field notes

- Host / container `_sum` fields are often **host-hours** for the month. The script converts them to average concurrent counts with ÷ 720.
- Ingested spans often come from `twol_ingested_events_bytes_sum` (Tracing Without Limits). `apm_ingest_gb_sum` is frequently null when usage is within Datadog’s included APM ingest allocation.
- Indexed spans come from `trace_search_indexed_events_count_sum`.

---

## Conversion assumptions

Defaults live at the top of `dd_usage_pull.py` and can be changed there.

| Parameter | Default | Meaning |
|---|---|---|
| `AVG_LOG_SIZE_KB` | 2.5 | Assumed avg log line size |
| `AVG_SPAN_SIZE_KB` | 1.5 | Assumed avg span size |
| `TS_PER_HOST` | 750 | Time series per host/container |
| `TS_TO_UNITS` | 3.3e-5 | NumSeries → Coralogix units/day |
| `DAYS_PER_MONTH` | 30 | Monthly → daily |
| `HOURS_PER_MONTH` | 720 | Host-hours → concurrent hosts |
| Log tier split | 70% / 30% | Monitoring / Compliance |
| Span tier split | 10% / 90% | Monitoring / Compliance |

### Logs

```text
daily_logs_gb = (ingested_logs_bytes + security_logs_bytes) / 1e9 / 30
monitoring    = daily_logs_gb × 0.70
compliance    = daily_logs_gb × 0.30
```

### Metrics

```text
total_ts = (hosts + containers) × 750
         + custom_metrics
         + serverless_functions_daily × 0.30
         + serverless_invocations_daily × 0.30
```

### Tracing

```text
daily_spans_gb = twol_ingested_events_bytes / 1e9 / 30
monitoring     = daily_spans_gb × 0.10
compliance     = daily_spans_gb × 0.90
```

### RUM

```text
sessions_day   = rum_total_sessions / 30
recordings_day = session_replay / 30
```

### Synthetics → Checkly

```text
api_checks_day     = synthetics_api_test_runs / 30
browser_checks_day = synthetics_browser_test_runs / 30
```

---

## Troubleshooting

| Error / symptom | Likely cause | Fix |
|---|---|---|
| `DD_API_KEY is not set` | Missing credential | Add keys to `.env` |
| `403 Forbidden` | Missing `usage_read` | Fix App/API key scopes |
| `400 Bad Request` | Bad month format | Use `YYYY-MM` |
| Metric shows 0 unexpectedly | Field name / free-tier null | Inspect `datadog_raw_*.json` in the ZIP |
| Billable summary skipped | N/A — not collected | Script only uses usage/summary |
| Month missing from window | No usage returned for that month | Check account activity / permissions |

---

## Permissions checklist

- [x] API key: `usage_read`
- [x] App key: `usage_read`

`billing_read` is **not** required or used.

For multi-org accounts, use a **parent organization** key if you need org-wide totals.
