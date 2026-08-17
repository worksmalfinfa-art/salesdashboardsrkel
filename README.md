# GROVE Sales Analytics

Area management & sales intelligence dashboard for **GROVE at CIBIS** — a mixed-use commercial lifestyle property in Jakarta.

**Live app:** https://grove-sales-analytics.streamlit.app/
**Version:** 2.0.0-cloud

---

## Overview

Consolidates sales data from two POS systems into role-based dashboards with multi-format export:

- **F&B tenants** — ESB POS (`.xlsx` exports)
- **Twist N' Turns Playground** — separate CSV-based POS

## Tech stack

| Component | Technology |
|---|---|
| Frontend + backend | Streamlit 1.45.1 (single `app.py`) |
| Database | **Supabase (PostgreSQL)** via `supabase-py` 2.x |
| Charts | Plotly 6.0.1 |
| Auth | Google OAuth (Streamlit viewer auth) + email fallback |
| Hosting | Streamlit Community Cloud (Python 3.11 via `runtime.txt`) |
| XLSX export | openpyxl 3.1.5 native charts |
| PDF export | Self-contained HTML + Plotly.js CDN, print via browser |

## Repository layout

```
app.py                          Main application (single file)
requirements.txt                Python dependencies
runtime.txt                     Pins python-3.11.0 for Streamlit Cloud
supabase/schema.sql             Database schema — run once in Supabase Studio
.streamlit/config.toml          GROVE green theme
.streamlit/secrets.toml.example Secrets template (real secrets are gitignored)
.gitignore                      Excludes secrets, credentials, data
README.md                       This file
```

## Database schema

PostgreSQL, defined in [`supabase/schema.sql`](supabase/schema.sql). Run it once in **Supabase Studio → SQL Editor**; every statement is idempotent, so re-running is safe.

| Table | Purpose | Key |
|---|---|---|
| `users` | Auth & RBAC | `email` (PK) |
| `tenants` | Tenant registry | `tenant_id` (PK), `tenant_name` and `esb_branch_name` unique |
| `sales_data` | F&B hourly sales from ESB POS | FK → `tenants(tenant_name)`, unique on (tenant, date, hour) |
| `playground_sales` | Playground TnT transactions | `order_id` (PK) |
| `upload_log` | Upload audit trail | `upload_id` (PK) |

Two constraints do work the application used to do by hand:

- **`sales_data` FK with `ON UPDATE`/`ON DELETE CASCADE`** — renaming or deleting a tenant now rewrites or removes its sales rows automatically. The Sheets version looped over every affected row one API call at a time.
- **`sales_data` unique (tenant, date, hour)** and **`playground_sales.order_id` PK** — de-duplication is enforced by the database, so a double upload cannot create duplicate rows even if the UI check is bypassed.

Row Level Security is enabled on all five tables with no permissive policies. The app connects with the `service_role` key, which bypasses RLS; access control is enforced in `check_auth()` via `users.role` and `users.tenant_access`.

## Pages

| Menu | Function | Description |
|---|---|---|
| 🏠 Master Dashboard | `page_master_dashboard()` | Consolidated F&B + Playground (3 tabs) |
| 🍽️ Dashboard F&B | `page_dashboard_fnb()` | ESB tenant analytics (7 tabs) |
| 🎪 Dashboard Playground | `page_dashboard_playground()` | Playground TnT (6 tabs) |
| 📤 Upload F&B (ESB) | `page_upload_esb()` | Parse & upload ESB `.xlsx` |
| 📤 Upload Playground | `page_upload_playground()` | Parse & upload Playground `.csv` |
| 🏢 Kelola Tenant | `page_tenants()` | Tenant CRUD (Admin+) |
| 👥 Kelola User | `page_users()` | User & role management (Super Admin) |
| 📋 Upload Log | `page_upload_log()` | Upload audit trail |

Each dashboard exports to **HTML (print/PDF)**, **XLSX with native charts**, **raw CSV**, and **summary CSV**.

## Setup

1. Create a Supabase project.
2. Run [`supabase/schema.sql`](supabase/schema.sql) in **SQL Editor**.
3. Edit the seed `INSERT` at the bottom of that file so each `esb_branch_name` matches the exact `Branch Name` string in that tenant's ESB export — see the warning below.
4. Copy `url` and the **`service_role`** key from **Project Settings → Data API** into Streamlit Cloud → **App settings → Secrets**, in the format shown in [`.streamlit/secrets.toml.example`](.streamlit/secrets.toml.example).

Access is restricted to the `@srkel.id` and `@teamup.id` domains (`ALLOWED_DOMAINS` in `app.py`).

### Local development

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # then fill in real values
streamlit run app.py
```

## Data source formats

**ESB POS (`.xlsx`)** — the **branch/hour summary** report. Headers at row 14, data from row 15. Columns: Branch Name (1), Sales Type (2), Sales Time (3), Sales Hour (4), Bill Total (5), Pax Total (6), Subtotal (7), Discount Total (8), Menu Discount Total (9), Nett Sales Total (10). Parsed by `ESBParser`.

> ⚠️ **Not every ESB export works.** The "Export Data ESB" report is a *per-item transaction* export with a different layout (metadata rows 1–8, headers at row 10) and will fail on upload. Quick check before uploading: **cell A14 must read `Branch Name`.**

> ⚠️ **`esb_branch_name` must match exactly.** The import joins ESB rows to tenants on this string. If it differs by so much as a suffix, the upload finds no matching tenant and imports nothing.

**Playground POS (`.csv`)** — flat CSV, per-transaction granularity. Columns: Order ID, Amount, Nett Sales, Tax Amount, Name, Date, Child Total, Companion Total. Parsed by `PlaygroundParser`; de-duplicated by `order_id` in the database.

## Notes & constraints

- **PostgREST caps responses at 1000 rows.** `_fetch_table()` pages through explicitly. Removing that paging would silently truncate every table at its first 1000 rows — a bug that looks like missing data, not an error.
- **The `service_role` key bypasses RLS.** It belongs only in Streamlit secrets — never in Git, never sent to a browser.
- **Reads are cached** with `@st.cache_data(ttl=300)` and cleared on every write via `_invalidate()`. This is a responsiveness measure for Streamlit's rerun-on-every-interaction model, not a quota workaround.
- **Plotly 6.x** — hex colors need the `#` prefix; `nbins` rejects `numpy.int64` (cast with `int()`); `height` cannot be passed twice alongside `**chart_layout`.
- **App sleep** — Streamlit Cloud free tier sleeps after 7 days idle and wakes in ~30s.

## Migrated from Google Sheets

Versions before this used Google Sheets as the database. The move to Postgres replaced the `SheetsDB` class only — all parsers, dashboards, charts and exports were untouched, and the public data-layer API is unchanged (19 methods, same names and signatures).

What it resolved:

| Was | Now |
|---|---|
| HTTP 429 after a few page navigations (60 reads/min quota) | No quota |
| ~80 lines of hand-rolled `session_state` caching to stay under that quota | ~10 lines of `st.cache_data` |
| Degrades past ~50K rows | Indexed on `sales_date` and `(tenant_name, sales_date)` |
| Tenant rename rewrote each sales row via its own API call | One `UPDATE`, cascaded by the FK |
| De-duplication enforced only in Python | Enforced by database constraints |
