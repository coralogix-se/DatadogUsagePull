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
- Datadog **API key** from the org you want to measure (API keys have no scopes)
- Datadog **Application key** whose owner (user or service account) has the
  **Usage Read** permission (`usage_read`) — by default only the Datadog Admin
  role has it. If the App key is scoped, include `usage_read` in its scopes.
- Both keys must come from a **parent-level org** (a standalone org counts).
  `/api/v1/usage/summary` is not available to child orgs.

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
DD_SITE=datadoghq.com        # the site you log in to — see "Datadog sites" below
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

### Datadog sites

Use the site that matches the URL you log in to. A pasted login URL
(e.g. `https://app.datadoghq.eu`) is also accepted.

| You log in at | `DD_SITE` |
|---|---|
| `app.datadoghq.com` | `datadoghq.com` (US1, default) |
| `us3.datadoghq.com` | `us3.datadoghq.com` |
| `us5.datadoghq.com` | `us5.datadoghq.com` |
| `app.datadoghq.eu` | `datadoghq.eu` |
| `ap1.datadoghq.com` | `ap1.datadoghq.com` |
| `ap2.datadoghq.com` | `ap2.datadoghq.com` |
| `app.ddog-gov.com` | `ddog-gov.com` (US1-FED) |

Keys only work on their own site — a wrong `DD_SITE` returns **403**, not 401.

---

## Datadog API endpoints used

Official [Usage Metering API](https://docs.datadoghq.com/api/latest/usage-metering/) only — no UI scraping, no cost endpoints.

| Endpoint | Purpose |
|---|---|
| `GET /api/v1/usage/summary` | Product usage volumes for the 3-month window |
| `GET /api/v2/usage/hourly_usage?filter[product_families]=timeseries` | Custom-metrics fallback — only called for months where the summary reports zero custom metrics |

The script does **not** call `estimated_cost`, `historical_cost`, `projected_cost`, or `billable-summary`.

> Datadog usage data can lag up to ~72 hours.

### Important field notes

- Host / container `_sum` fields are often **host-hours** for the month. The script converts them to average concurrent counts with ÷ 720.
- Ingested spans often come from `twol_ingested_events_bytes_sum` (Tracing Without Limits). `apm_ingest_gb_sum` is frequently null when usage is within Datadog’s included APM ingest allocation.
- Indexed spans come from `trace_search_indexed_events_count_sum`.
- Custom-metric fields in `/v1/usage/summary` (`custom_ts_avg` etc.) are often null/0 even when the account has custom metrics; the v2 hourly usage `timeseries` family is the authoritative source and is used as a fallback.

---

## Troubleshooting

| Error / symptom | Likely cause | Fix |
|---|---|---|
| `DD_API_KEY is not set` | Missing credential | Add keys to `.env` |
| `403` — script says *API key is NOT valid on site* | Wrong `DD_SITE`, or bad/revoked API key | Set `DD_SITE` to the site you log in to; re-copy the API key |
| `403` — API key valid on site | App key owner lacks **Usage Read**; scoped App key missing `usage_read`; keys from a **child org**; API/App keys from different orgs; org IP allowlist | Create the App key as an Admin (or a service account with Usage Read) in the **parent** org; check IP allowlist |
| `400 Bad Request` | Bad month format | Use `YYYY-MM` |
| Metric shows 0 unexpectedly | Field name / free-tier null | Inspect `datadog_raw_*.json` in the ZIP |
| Billable summary skipped | N/A — not collected | Script only uses usage/summary |
| Month missing from window | No usage returned for that month | Check account activity / permissions |

---

## Permissions checklist

- [x] API key: from the parent org (API keys have no scopes)
- [x] App key: owner has **Usage Read**; `usage_read` in scopes if the key is scoped
- [x] Correct `DD_SITE` for the org

Quick self-check (replace the host with your site's API host):

```bash
curl -s "https://api.datadoghq.com/api/v1/validate" -H "DD-API-KEY: $DD_API_KEY"
curl -s "https://api.datadoghq.com/api/v1/usage/summary?start_month=2026-01" \
  -H "DD-API-KEY: $DD_API_KEY" -H "DD-APPLICATION-KEY: $DD_APP_KEY"
```

`billing_read` is **not** required or used.

For multi-org accounts, keys **must** come from the **parent organization** — child-org keys get 403 on the usage summary endpoint.
