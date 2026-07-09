# Datadog → Coralogix Usage Report

A self-contained Python script that pulls your Datadog usage from the official
Usage Metering API, mirrors every tile on the Bill Overview page, converts the
numbers to Coralogix sizing estimates, and packages everything into a ZIP file.

---

## What you get

Running the script produces a single ZIP file containing:

| File | Contents |
|---|---|
| `datadog_raw_YYYY-MM.json` | Full raw API responses (for reference / audit) |
| `datadog_usage_YYYY-MM.csv` | All Bill Overview metrics in a flat CSV |
| `coralogix_sizing_YYYY-MM.xlsx` | Excel workbook: Bill Overview + Coralogix sizing + billable breakdown |
| `report_YYYY-MM.html` | Self-contained HTML report — open in any browser |

Send the ZIP back to your Coralogix contact. The HTML report shows the
four numbers to enter into the Coralogix TCO Calculator:

- **Logs** — GB / day
- **Metrics** — NumSeries (TimeSeries)
- **Tracing** — GB / day
- **RUM** — sessions / month

---

## Requirements

- Python 3.9 or later
- A Datadog **API key** with `usage_read` permission
- A Datadog **Application key** with `usage_read` and (optionally) `billing_read` permission

> Cost data (estimated / historical / projected) requires `billing_read` on a
> parent-organization key. If that permission is missing the script still runs
> and skips only the cost section.

---

## Quick start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

If you are in a managed Python environment (e.g. macOS system Python or a
Homebrew installation), use a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Configure

Copy the example config and fill in your Datadog credentials:

```bash
cp .env.example .env
```

Open `.env` in any text editor and set:

```
DD_API_KEY=<your API key>
DD_APP_KEY=<your Application key>
DD_SITE=datadoghq.com        # or datadoghq.eu, us3.datadoghq.com, etc.
DD_MONTH=2026-06             # optional — defaults to previous month
```

### 3. Run

```bash
python dd_usage_pull.py
```

The script prints a live progress log and, when finished, shows a summary:

```
  Datadog → Coralogix Usage Report
  Site  : datadoghq.com
  Month : 2026-06
  Output: /path/to/current/directory

  [1/5] Fetching usage summary …        OK
  [2/5] Fetching billable usage summary … OK
  [3/5] Fetching estimated cost …       OK
  [4/5] Fetching historical cost …      OK
  [5/5] Fetching projected cost …       OK

  ══════════════════════════════════════════
  Coralogix TCO Sizing Summary — 2026-06
  ══════════════════════════════════════════
  Logs     : 166.67 GB/day
  Metrics  : 4,115,870 NumSeries
  Tracing  : 80.00 GB/day
  RUM      : 200,000 sessions/month
  Est. Cost: $42,500.00 USD (MTD)
  ══════════════════════════════════════════

  Output ZIP : /path/to/datadog_coralogix_report_2026-06.zip
```

---

## CLI options

You can also pass arguments directly instead of using a `.env` file:

```bash
python dd_usage_pull.py --month 2026-05 --site datadoghq.eu --out /tmp/reports
```

| Option | Default | Description |
|---|---|---|
| `--month` | previous month | Target month in `YYYY-MM` format |
| `--site` | `datadoghq.com` | Datadog site |
| `--out` | `.` (current directory) | Output directory for the ZIP |

Environment variables (`DD_API_KEY`, `DD_APP_KEY`, `DD_SITE`, `DD_MONTH`) take
precedence over `.env` file values; CLI flags take precedence over environment
variables.

---

## Datadog API endpoints used

All calls go to the official [Usage Metering API](https://docs.datadoghq.com/api/latest/usage-metering/).
No UI scraping. No browser automation.

| Endpoint | Purpose |
|---|---|
| `GET /api/v1/usage/summary` | All product usage totals for the month |
| `GET /api/v1/usage/billable-summary` | Committed vs on-demand split per dimension |
| `GET /api/v2/usage/estimated_cost` | Month-to-date estimated cost |
| `GET /api/v2/usage/historical_cost` | Closed-month historical cost |
| `GET /api/v2/usage/projected_cost` | End-of-month cost forecast |

> **Data freshness:** Datadog usage data may lag up to 72 hours.
> Historical cost for a closed month is available by the 16th of the following month.

---

## Coralogix sizing formulas

The conversion uses the same formulas as the Datadog → Coralogix sizing
Excel template. All assumptions are visible at the top of `dd_usage_pull.py`
and can be changed there if needed.

| Parameter | Default | Effect |
|---|---|---|
| `AVG_LOG_SIZE_KB` | 2.5 KB | Average compressed log line size |
| `AVG_SPAN_SIZE_KB` | 1.5 KB | Average span payload size |
| `TS_PER_HOST` | 750 | Time series generated per host or container |
| `TS_TO_UNITS` | 3.3 × 10⁻⁵ | Coralogix metrics Units/day per TimeSeries |
| `DAYS_PER_MONTH` | 30 | Used for monthly → daily conversions |
| Log tier split | 50 / 40 / 10 % | Frequent Search / Monitoring / Compliance |

### Logs

```
total_ingested_gb   = ingested_logs_gb + security_logs_gb
total_indexed_count = (live_3d + live_7d + live + rehydrated + 15d + 30d+ retention)
indexed_size_gb     = total_indexed_count × avg_log_size_kb × 1024 ÷ 1024³
daily_gb            = total_ingested_gb ÷ 30
```

### Metrics

```
total_ts = (infra_hosts + apm_hosts + profiled_hosts + fargate_types + containers) × 750
         + custom_metrics
         + (daily_serverless_functions × 0.30)
         + (daily_serverless_invocations × 0.30)
units_per_day = total_ts × 3.3e-5
```

### Tracing

```
ingested_spans_gb  = ingested_spans_bytes ÷ 1e9
indexed_spans_gb   = (indexed_spans + custom_events) × avg_span_size_kb ÷ 1024²
daily_ingest_gb    = ingested_spans_gb ÷ 30
```

### RUM

```
rum_sessions_monthly = rum_total_session_count
```

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `DD_API_KEY is not set` | Missing credential | Add key to `.env` |
| `403 Forbidden` | Missing permission | Ensure `usage_read` on API key; `billing_read` on App key |
| `400 Bad Request` | Invalid month format | Use `YYYY-MM` format, e.g. `2026-06` |
| All metrics show 0 | Empty API response | Check that the account has usage in the target month |
| Cost section missing | `billing_read` not on App key | Normal — usage data still exported |

If a metric shows 0 and you expect a value, check `datadog_raw_YYYY-MM.json`
inside the ZIP — it contains the complete raw API responses for inspection.

---

## Permissions checklist

Create a dedicated Datadog Application key with:

- [x] `usage_read` — required for all usage data
- [x] `billing_read` — required for cost data (estimated, historical, projected)

The API key only needs `usage_read`.

For multi-org accounts, the key must belong to the **parent organization** to
see cross-org billable and cost data.
