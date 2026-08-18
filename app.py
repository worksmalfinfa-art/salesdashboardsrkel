"""
GROVE Sales Analytics — CLOUD Edition
======================================
Area Management Tool for GROVE at CIBIS
Cloud deployment — Supabase (PostgreSQL) + Streamlit Cloud + Google OAuth

Tech Stack: Streamlit Cloud + Supabase + Plotly + Google OAuth
Version: 2.0.0-cloud
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from plotly.subplots import make_subplots

# Set global plotly font size
# Every chart in the app inherits this. Transparent ground so a chart sits on
# its card instead of on a white slab of its own; horizontal hairlines only,
# since vertical rules add nothing to a time series; rounded bar caps to match
# the card system; and muted axis type so the data outranks the furniture.
pio.templates["grove"] = go.layout.Template(
    layout=go.Layout(
        colorway=["#1A6B3F", "#C08A2C", "#6B7B3A", "#3B4C7A", "#8A5B6E",
                  "#7A5230", "#2E9160", "#9E6B4A"],
        font=dict(family="ui-rounded, Segoe UI, system-ui, sans-serif",
                  size=12, color="#4A524D"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        # A share of the bar width, not a fixed pixel count: 8px turned the thin
        # bars of a 30-day series into capsules while leaving wide bars barely
        # rounded. A percentage behaves at both extremes.
        barcornerradius="10%",
        # The title owns the top margin, the legend the bottom one. Both used to
        # sit at x=0 directly above the plot, so any chart with a legend drew
        # its title straight through the swatches.
        # Titles now live above the chart as HTML headings, so the top
        # margin only has to clear the legend.
        margin=dict(t=34, b=48, l=54, r=18),
        title=dict(font=dict(size=13.5, color="#131715"), x=0, xanchor="left",
                   y=0.98, yanchor="top"),
        xaxis=dict(showgrid=False, zeroline=False, linecolor="#EBECE8",
                   tickfont=dict(size=11, color="#8B948D"),
                   title=dict(font=dict(size=11, color="#8B948D"))),
        yaxis=dict(showgrid=True, gridcolor="#F1F2EF", zeroline=False,
                   linecolor="rgba(0,0,0,0)",
                   tickfont=dict(size=11, color="#8B948D"),
                   title=dict(font=dict(size=11, color="#8B948D"))),
        # Title left, legend right, both in the top margin. Separating them
        # horizontally rather than vertically is what keeps them apart at any
        # chart height -- legend y is a fraction of the plot area, so a bottom
        # legend that clears the axis on a 300px chart is pushed past the
        # margin and clipped on a 500px one.
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=1, xanchor="right",
                    font=dict(size=11), bgcolor="rgba(0,0,0,0)",
                    itemsizing="constant"),
        hoverlabel=dict(bgcolor="#131715", bordercolor="#131715",
                        font=dict(size=12, color="#FFFFFF",
                                  family="ui-rounded, Segoe UI, system-ui, sans-serif")),
        hovermode="closest",
    ))
pio.templates.default = "plotly_white+grove"
from supabase import create_client
import os
import re
import json
from datetime import datetime, timedelta, date
from io import BytesIO
import openpyxl

pd.set_option("future.no_silent_downcasting", True)

# ============================================================================
# CONFIGURATION
# ============================================================================
APP_TITLE = "GROVE Sales Analytics"
APP_ICON = "📊"
APP_VERSION = "2.0.0-cloud"

TIME_SEGMENTS = {
    "Breakfast": (7, 10),
    "Lunch": (12, 14),
    "After Office": (17, 19),
}

ESB_HEADER_ROW = 14
ESB_DATA_START_ROW = 15
ESB_COL_MAP = {
    "branch_name": 1, "sales_type": 2, "sales_date": 3, "sales_hour": 4,
    "bill_total": 5, "pax_total": 6, "subtotal": 7,
    "discount_total": 8, "menu_discount": 9, "nett_sales": 10,
}

COLORS = {
    "primary": "#103D28", "secondary": "#1A6B3F", "accent": "#2E9160",
    "gold": "#C08A2C", "dark": "#0B140F", "light": "#DCEEE4",
    "red": "#C0483C", "blue": "#3B4C7A", "orange": "#9E6B4A", "white": "#FFFFFF",
}
# One palette, defined once. Charts and tenant marks draw from the same list
# so a brand keeps its hue whether it appears as a bar, a line or a square.
TENANT_HUES = ["#1A6B3F", "#C08A2C", "#6B7B3A", "#3B4C7A", "#8A5B6E", "#7A5230",
               "#2E9160", "#9E6B4A", "#4A6B7B", "#7B4A6B"]
CHART_PALETTE = TENANT_HUES

ALLOWED_DOMAINS = ["srkel.id", "teamup.id"]


# ============================================================================
# Supabase (PostgreSQL) DATABASE LAYER
# ----------------------------------------------------------------------------
# Public API is identical to the Google Sheets layer this replaces, so the
# parsers, dashboards, uploads and exports call it unchanged.
#
# The hand-rolled session_state cache is gone. It existed to stay under the
# Sheets quota of 60 reads/min; Postgres has no such quota, so the only thing
# still worth avoiding is refetching a whole table on every Streamlit rerun.
# st.cache_data handles that in a few lines.
# ============================================================================

# PostgREST caps every response at 1000 rows no matter what limit is requested,
# so full-table reads must page explicitly. Without this, any table past 1000
# rows would silently load only its first page.
_PAGE_SIZE = 1000
# Rows per insert request — keeps each JSON payload comfortably small.
_CHUNK_SIZE = 500


@st.cache_resource(show_spinner=False)
def _get_client():
    """One Supabase client per server process."""
    cfg = st.secrets["supabase"]
    return create_client(cfg["url"], cfg["service_key"])


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_table(table, order_col=None, descending=False):
    """Read an entire table into a DataFrame, paging past the 1000-row cap."""
    client = _get_client()
    rows, start = [], 0
    while True:
        q = client.table(table).select("*")
        if order_col:
            q = q.order(order_col, desc=descending)
        page = q.range(start, start + _PAGE_SIZE - 1).execute().data or []
        rows.extend(page)
        if len(page) < _PAGE_SIZE:
            break
        start += _PAGE_SIZE
    return pd.DataFrame(rows) if rows else pd.DataFrame()


class SupabaseDB:
    _SALES_COLS = ("tenant_id", "sales_date", "sales_hour", "pax_total",
                   "subtotal", "discount_total", "nett_sales",
                   "uploaded_by", "uploaded_at")
    _PG_COLS = ("order_id", "amount", "nett_sales", "tax_amount",
                "customer_name", "sales_date", "child_total",
                "companion_total", "uploaded_by", "uploaded_at")

    def __init__(self):
        self.client = _get_client()

    # --- internals ---
    @staticmethod
    def _invalidate():
        """Drop cached reads after a write."""
        _fetch_table.clear()

    def _count(self, table):
        # head=True asks PostgREST for the count alone, with no rows in the body.
        res = self.client.table(table).select("*", count="exact", head=True).execute()
        return res.count or 0

    def _insert_chunked(self, table, records, on_conflict=None, ignore_duplicates=False):
        if not records:
            return
        for i in range(0, len(records), _CHUNK_SIZE):
            chunk = records[i:i + _CHUNK_SIZE]
            q = self.client.table(table)
            if on_conflict:
                q.upsert(chunk, on_conflict=on_conflict,
                         ignore_duplicates=ignore_duplicates).execute()
            else:
                q.insert(chunk).execute()
        self._invalidate()

    # --- User Management ---
    def get_user(self, email):
        df = _fetch_table("users")
        if df.empty: return None
        match = df[df["email"] == email]
        return match.iloc[0].to_dict() if not match.empty else None

    def create_user(self, email, display_name, role, tenant_access, created_by):
        self.client.table("users").insert({
            "email": email, "display_name": display_name, "role": role,
            "tenant_access": tenant_access, "created_by": created_by,
        }).execute()
        self._invalidate()

    def get_all_users(self):
        return _fetch_table("users", order_col="created_at")

    def has_users(self):
        return self._count("users") > 0

    def update_user_status(self, email, is_active):
        val = is_active if isinstance(is_active, bool) \
              else str(is_active).upper() in ("TRUE", "1")
        self.client.table("users").update({"is_active": val}).eq("email", email).execute()
        self._invalidate()

    def update_user_role(self, email, new_role, new_access):
        self.client.table("users").update({
            "role": new_role, "tenant_access": new_access,
        }).eq("email", email).execute()
        self._invalidate()

    # --- Tenants (the brand) ---
    def get_tenants(self, active_only=False):
        df = _fetch_table("tenants", order_col="tenant_id")
        if active_only and not df.empty and "is_active" in df.columns:
            df = df[df["is_active"] == True]
        return df

    def _next_id(self, df, col, prefix):
        """Lowest unused id, so deleting then adding cannot collide."""
        used = set(df[col].astype(str)) if not df.empty and col in df.columns else set()
        n = 1
        while f"{prefix}{n:03d}" in used:
            n += 1
        return f"{prefix}{n:03d}"

    def add_tenant(self, tenant_name, esb_branch_name=None, category="F&B", pos_type="esb"):
        tid = self._next_id(self.get_tenants(), "tenant_id", "T")
        self.client.table("tenants").insert({
            "tenant_id": tid, "tenant_name": tenant_name,
            "esb_branch_name": esb_branch_name or None,
            "category": category, "pos_type": pos_type,
        }).execute()
        self._invalidate()
        return tid

    def get_tenant_by_branch(self, esb_branch_name):
        df = self.get_tenants()
        if df.empty: return None
        match = df[df["esb_branch_name"] == esb_branch_name]
        return match.iloc[0].to_dict() if not match.empty else None

    def delete_tenant(self, tenant_id):
        res = self.client.table("sales_data").select("id", count="exact", head=True) \
                  .eq("tenant_id", tenant_id).execute()
        sales_count = res.count or 0
        out = self.client.table("tenants").delete().eq("tenant_id", tenant_id).execute()
        self._invalidate()
        return (1 if out.data else 0), sales_count

    def edit_tenant(self, tenant_id, new_name, new_branch,
                    category=None, pos_type=None, is_active=None):
        # sales_data keys on tenant_id now, so renaming a brand touches no
        # sales row at all -- the label lives in exactly one place.
        payload = {"tenant_name": new_name, "esb_branch_name": new_branch or None}
        if category is not None:  payload["category"] = category
        if pos_type is not None:  payload["pos_type"] = pos_type
        if is_active is not None: payload["is_active"] = bool(is_active)
        res = self.client.table("tenants").update(payload).eq("tenant_id", tenant_id).execute()
        self._invalidate()
        return bool(res.data)

    # --- Units (the physical space) ---
    def get_units(self):
        return _fetch_table("units", order_col="unit_id")

    def add_unit(self, unit_code, unit_name=None, floor=None, area_sqm=None):
        uid = self._next_id(self.get_units(), "unit_id", "U")
        self.client.table("units").insert({
            "unit_id": uid, "unit_code": unit_code, "unit_name": unit_name,
            "floor": floor, "area_sqm": area_sqm,
        }).execute()
        self._invalidate()
        return uid

    def edit_unit(self, unit_id, **fields):
        allowed = {k: v for k, v in fields.items()
                   if k in ("unit_code", "unit_name", "floor", "area_sqm", "is_active")}
        if not allowed: return False
        res = self.client.table("units").update(allowed).eq("unit_id", unit_id).execute()
        self._invalidate()
        return bool(res.data)

    def delete_unit(self, unit_id):
        res = self.client.table("units").delete().eq("unit_id", unit_id).execute()
        self._invalidate()
        return bool(res.data)

    # --- Tenancies (who occupied which unit, when) ---
    def get_tenancies(self):
        return _fetch_table("tenancies", order_col="start_date", descending=True)

    def get_tenancies_detailed(self):
        """Tenancies with unit code and brand name attached, for display."""
        tn = self.get_tenancies()
        if tn.empty: return tn
        out, units, tenants = tn.copy(), self.get_units(), self.get_tenants()
        if not units.empty:
            out = out.merge(units[["unit_id", "unit_code"]], on="unit_id", how="left")
        if not tenants.empty:
            out = out.merge(tenants[["tenant_id", "tenant_name"]], on="tenant_id", how="left")
        out["status"] = out["end_date"].isna().map({True: "Aktif", False: "Selesai"})
        return out

    def get_active_tenancy(self, unit_id):
        df = self.get_tenancies()
        if df.empty: return None
        m = df[(df["unit_id"] == unit_id) & (df["end_date"].isna())]
        return m.iloc[0].to_dict() if not m.empty else None

    def add_tenancy(self, unit_id, tenant_id, start_date, end_date=None, notes=None):
        self.client.table("tenancies").insert({
            "unit_id": unit_id, "tenant_id": tenant_id,
            "start_date": str(start_date),
            "end_date": str(end_date) if end_date else None,
            "notes": notes,
        }).execute()
        self._invalidate()

    def end_tenancy(self, tenancy_id, end_date):
        self.client.table("tenancies").update({"end_date": str(end_date)}) \
            .eq("tenancy_id", int(tenancy_id)).execute()
        self._invalidate()

    def replace_tenant(self, unit_id, new_tenant_id, handover_date):
        """
        Hand a unit over to a different brand.

        Closes the outgoing tenancy the day before the hand-over and opens the
        incoming one on it, so the unit is never occupied twice on the same day
        -- which the database would reject anyway. No sales row is touched:
        each brand keeps exactly what it earned, before and after.
        """
        handover = pd.to_datetime(handover_date).date()
        current = self.get_active_tenancy(unit_id)
        if current:
            started = pd.to_datetime(current["start_date"]).date()
            if started >= handover:
                raise ValueError(
                    "Tanggal serah terima harus setelah tanggal mulai penyewa "
                    f"saat ini ({started:%d %b %Y})."
                )
            self.end_tenancy(current["tenancy_id"], handover - timedelta(days=1))
        self.add_tenancy(unit_id, new_tenant_id, handover,
                         notes=f"Serah terima {handover:%Y-%m-%d}")

    # --- Targets ---
    def get_targets(self):
        return _fetch_table("targets", order_col="period_month", descending=True)

    def upsert_target(self, tenant_id, period_month, target_nett, updated_by=None):
        self.client.table("targets").upsert({
            "tenant_id": tenant_id,
            "period_month": str(period_month),
            "target_nett": int(target_nett),
            "updated_by": updated_by,
        }, on_conflict="tenant_id,period_month").execute()
        self._invalidate()

    # --- Sales Data ---
    def get_sales_data(self, tenant_filter=None):
        # Read the view, not the table: it resolves which unit each sale
        # belonged to by matching the sale's date against the tenancy periods,
        # so a hand-over splits history at the right day without the app
        # having to get that date logic right in every dashboard.
        df = _fetch_table("v_sales_enriched", order_col="sales_date")
        if df.empty: return df
        df = df.copy()
        df["sales_date"] = pd.to_datetime(df["sales_date"], errors="coerce")
        for col in ["pax_total", "subtotal", "discount_total", "nett_sales"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        if tenant_filter and tenant_filter != "All":
            df = df[df["tenant_name"] == tenant_filter]
        return df

    def append_sales_data(self, rows):
        records = [dict(zip(self._SALES_COLS, r)) for r in rows]
        # Collapse repeats of the same slot inside one file before sending.
        # Postgres refuses an upsert that touches the same key twice ("ON
        # CONFLICT DO UPDATE command cannot affect row a second time"), and an
        # ESB export can legitimately split one hour over several rows when the
        # branch records more than one sales type. Those are separate real
        # sales in the same hour, so they add up rather than overwrite.
        merged = {}
        for rec in records:
            key = (rec["tenant_id"], rec["sales_date"], rec["sales_hour"])
            if key in merged:
                for f in ("pax_total", "subtotal", "discount_total", "nett_sales"):
                    merged[key][f] += rec[f]
            else:
                merged[key] = dict(rec)
        # Upsert on the (tenant, date, hour) unique constraint. Re-uploading the
        # same file therefore rewrites those slots with identical values instead
        # of appending a second copy -- the behaviour the Sheets version lacked,
        # where every upload simply appended and a repeat doubled the day.
        self._insert_chunked("sales_data", list(merged.values()),
                             on_conflict="tenant_id,sales_date,sales_hour")

    def count_existing_slots(self, tenant_id, rows):
        """How many (date, hour) slots in this file the tenant already has.
        Lets the upload screen say what will be overwritten before it happens."""
        df = _fetch_table("sales_data")
        if df.empty or "tenant_id" not in df.columns: return 0
        have = {(str(r["sales_date"])[:10], str(r["sales_hour"]))
                for _, r in df[df["tenant_id"] == tenant_id].iterrows()}
        incoming = {(str(r["sales_date"])[:10], str(r["sales_hour"])) for r in rows}
        return len(have & incoming)

    def get_existing_keys(self, tenant_id):
        df = _fetch_table("sales_data")
        if df.empty or "tenant_id" not in df.columns: return set()
        df = df[df["tenant_id"] == tenant_id]
        return {(str(r["tenant_id"]), str(r["sales_date"])[:10], str(r["sales_hour"]))
                for _, r in df.iterrows()}

    # --- Playground Data ---
    def get_playground_data(self):
        df = _fetch_table("playground_sales", order_col="sales_date")
        if df.empty: return df
        df = df.copy()
        df["sales_date"] = pd.to_datetime(df["sales_date"], errors="coerce")
        for col in ["amount", "nett_sales", "tax_amount", "child_total", "companion_total"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df

    def get_playground_tenant_id(self):
        """The tenant registered with pos_type 'playground', if any."""
        df = self.get_tenants()
        if df.empty or "pos_type" not in df.columns: return None
        m = df[df["pos_type"] == "playground"]
        return m.iloc[0]["tenant_id"] if not m.empty else None

    def insert_playground_batch(self, rows):
        records = [dict(zip(self._PG_COLS, r)) for r in rows]
        # Stamp the owning tenant so Playground is just another tenant to the
        # master dashboard rather than a hard-coded second world.
        pg_id = self.get_playground_tenant_id()
        if pg_id:
            for rec in records:
                rec["tenant_id"] = pg_id
        # order_id is the primary key, so the database rejects repeats. Collapse
        # duplicates inside the file first, since one statement cannot touch the
        # same key twice.
        seen, unique = set(), []
        for rec in records:
            oid = rec["order_id"]
            if oid in seen: continue
            seen.add(oid)
            unique.append(rec)
        before = self._count("playground_sales")
        self._insert_chunked("playground_sales", unique,
                             on_conflict="order_id", ignore_duplicates=True)
        return self._count("playground_sales") - before

    # --- Upload Log ---
    def log_upload(self, tenant_name, file_name, rows_added, uploaded_by):
        # Milliseconds included: upload_id is a primary key here, and two
        # uploads finishing in the same second would otherwise collide.
        uid = f"UPL-{datetime.now().strftime('%Y%m%d%H%M%S%f')[:-3]}"
        self.client.table("upload_log").insert({
            "upload_id": uid, "tenant_name": tenant_name, "file_name": file_name,
            "rows_added": rows_added, "uploaded_by": uploaded_by,
        }).execute()
        self._invalidate()

    def get_upload_log(self):
        return _fetch_table("upload_log", order_col="uploaded_at", descending=True)

    def get_db_stats(self):
        return {
            "sales_rows": self._count("sales_data"),
            "pg_rows": self._count("playground_sales"),
            "tenants": self._count("tenants"),
            "users": self._count("users"),
            "uploads": self._count("upload_log"),
            "db_size_kb": "supabase",
        }


# ============================================================================
# ESB FILE PARSER
# ============================================================================
class ESBParser:
    @staticmethod
    def parse(file_bytes, filename="upload.xlsx"):
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        ws = wb.active
        meta = {
            "company": ws.cell(2, 1).value or "",
            "period": ws.cell(5, 2).value or "",
        }
        rows = []
        for r in range(ESB_DATA_START_ROW, ws.max_row + 1):
            branch = ws.cell(r, ESB_COL_MAP["branch_name"]).value
            if not branch:
                continue
            raw_date = ws.cell(r, ESB_COL_MAP["sales_date"]).value
            if isinstance(raw_date, datetime):
                sales_date = raw_date.strftime("%Y-%m-%d")
            elif isinstance(raw_date, str):
                sales_date = raw_date[:10]
            else:
                continue
            sales_hour = str(ws.cell(r, ESB_COL_MAP["sales_hour"]).value or "")
            pax = ws.cell(r, ESB_COL_MAP["pax_total"]).value or 0
            subtotal = ws.cell(r, ESB_COL_MAP["subtotal"]).value or 0
            discount = (ws.cell(r, ESB_COL_MAP["discount_total"]).value or 0) + \
                       (ws.cell(r, ESB_COL_MAP["menu_discount"]).value or 0)
            nett = ws.cell(r, ESB_COL_MAP["nett_sales"]).value or 0
            rows.append({
                "branch_name": str(branch).strip(),
                "sales_date": sales_date,
                "sales_hour": sales_hour.strip(),
                "pax_total": int(pax), "subtotal": int(subtotal),
                "discount_total": int(discount), "nett_sales": int(nett),
            })
        return rows, meta


class PlaygroundParser:
    @staticmethod
    def parse(file_bytes, filename="upload.csv"):
        import csv as csvmod
        text = file_bytes.decode("utf-8-sig")
        reader = csvmod.DictReader(text.splitlines())
        rows = []
        for r in reader:
            try:
                rows.append({
                    "order_id": r.get("Order ID","").strip(),
                    "amount": int(float(r.get("Amount",0))),
                    "nett_sales": int(float(r.get("Nett Sales",0))),
                    "tax_amount": int(float(r.get("Tax Amount",0))),
                    "customer_name": r.get("Name","").strip().strip('"'),
                    "sales_date": r.get("Date","").strip(),
                    "child_total": int(float(r.get("Child Total",0))),
                    "companion_total": int(float(r.get("Companion Total",0))),
                })
            except (ValueError, TypeError):
                continue
        return rows


# ============================================================================
# STYLING
# ============================================================================
def apply_custom_css():
    """
    Colour, type and shape only.

    Everything structural -- containers, grids, spacing, heights -- is left to
    Streamlit. The previous stylesheet fought it: raw HTML was injected for
    cards, and Streamlit measures each markdown block in JavaScript without
    seeing padding added inside, so every card reported 16px shorter than it
    drew and sat on the element below. Patching that with guessed selectors and
    negative margins moved the damage around instead of removing it. No rule
    here sets position, height, or a negative margin.
    """
    st.markdown("""
    <style>
    :root{
      --bg:#EFEFEC; --card:#FFFFFF; --ink:#131715; --muted:#8B948D; --line:#E4E6E1;
      --g900:#103D28; --g700:#1A6B3F; --g500:#2E9160; --g50:#EEF6F1;
      --ochre:#C08A2C; --red:#C0483C;
    }
    .stApp{background:var(--bg)}
    html,body,[class*="css"],button,input,select{
      font-family:ui-rounded,"Segoe UI Variable Display","Segoe UI",system-ui,sans-serif}
    header[data-testid="stHeader"]{background:var(--bg)}

    /* Cards: Streamlit's own bordered container, restyled. Because Streamlit
       creates and measures it, it can never disagree with its contents. */
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"]){
      background:var(--card);border-color:var(--line);border-radius:16px}

    [data-testid="stMetricValue"]{
      font-size:1.7rem;font-weight:700;color:var(--ink);
      font-variant-numeric:tabular-nums;letter-spacing:-.02em}
    [data-testid="stMetricLabel"]{font-size:.8rem;color:var(--muted);font-weight:600}
    [data-testid="stMetricDelta"]{font-size:.78rem}

    h1,h2,h3,h4{color:var(--ink);letter-spacing:-.02em}
    h1{font-size:1.7rem!important;font-weight:800}
    h2{font-size:1.2rem!important;font-weight:700}
    h3{font-size:1rem!important;font-weight:700}
    .stCaption,[data-testid="stCaptionContainer"]{color:var(--muted)}

    section[data-testid="stSidebar"]{background:var(--card);border-right:1px solid var(--line)}
    section[data-testid="stSidebar"] [role="radiogroup"] label{
      padding:6px 10px;border-radius:10px;font-size:.86rem}
    section[data-testid="stSidebar"] [role="radiogroup"] label:hover{background:var(--bg)}
    section[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){
      background:var(--g50);font-weight:700}
    section[data-testid="stSidebar"] [role="radiogroup"] input{display:none}

    .stTabs [data-baseweb="tab-list"]{gap:2px;border-bottom:1px solid var(--line)}
    .stTabs [data-baseweb="tab"]{font-size:.86rem;font-weight:600;color:var(--muted)}
    .stTabs [aria-selected="true"]{color:var(--ink)!important}
    .stButton button,.stDownloadButton button{
      border-radius:9px;font-weight:600;font-size:.85rem;border:1px solid var(--line)}
    .stDataFrame{font-variant-numeric:tabular-nums}
    </style>
    """, unsafe_allow_html=True)


TENANT_HUES = ["#1A6B3F", "#C08A2C", "#6B7B3A", "#3B4C7A", "#8A5B6E", "#7A5230",
               "#2E9160", "#9E6B4A", "#4A6B7B", "#7B4A6B"]
CHART_PALETTE = TENANT_HUES
GREEN, GREEN_D, OCHRE, FLAT = "#1A6B3F", "#103D28", "#C08A2C", "#F0F1EE"


def tenant_hue(tenant_id, name=""):
    """Stable colour per tenant, keyed on id so a rename keeps the hue."""
    digits = "".join(c for c in str(tenant_id or "") if c.isdigit())
    key = int(digits) if digits else sum(ord(c) for c in str(name or tenant_id))
    return TENANT_HUES[key % len(TENANT_HUES)]


def fmt_rp(val):
    if val >= 1_000_000_000: return f"Rp {val/1_000_000_000:,.2f} M"
    if val >= 1_000_000:     return f"Rp {val/1_000_000:,.1f} Jt"
    return f"Rp {val:,.0f}"


def id_num(n, dec=0):
    return f"{n:,.{dec}f}".replace(",", "#").replace(".", ",").replace("#", ".")


def page_head(title, subtitle=None):
    st.header(title, anchor=False)
    if subtitle:
        st.caption(subtitle)


def card(title=None, subtitle=None):
    """
    A card is a bordered Streamlit container -- nothing else.

    Use as `with card("Title"):`. Because Streamlit builds and measures the
    container itself, its height always matches what is inside it.
    """
    c = st.container(border=True)
    if title:
        c.markdown(f"###### {title}")
    if subtitle:
        c.caption(subtitle)
    return c


def kpi(col, label, value, delta=None, help=None):
    """One metric in a bordered container, using st.metric rather than markup."""
    with col.container(border=True):
        st.metric(label, value, delta=delta, help=help)


DEFAULT_CHART_CONFIG = {
    "displayModeBar": True, "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d", "toImage",
                               "autoScale2d", "toggleSpikelines"],
    "scrollZoom": False}


def show_chart(fig, height=None, **kw):
    """
    Charts inherit the grove template; only the toolbar config is set here.

    Accepts and discards the keywords the old signature took, because 40 call
    sites still pass them and a rename that breaks a page at runtime is worse
    than a tolerant wrapper.
    """
    if height:
        fig.update_layout(height=height)
    kw.pop("use_container_width", None)
    cfg = kw.pop("config", None) or DEFAULT_CHART_CONFIG
    st.plotly_chart(fig, use_container_width=True, config=cfg, **kw)


def tenant_table(df, cols, height=None):
    """
    Leaderboards and breakdowns as a real dataframe.

    Hand-built rows of HTML were the single largest source of overlap; a
    dataframe is measured by Streamlit, sorts itself, and still carries
    progress bars and sparklines through column_config.
    """
    st.dataframe(df, use_container_width=True, hide_index=True,
                 height=height, column_config=cols)


# --- names the rest of the app still calls -------------------------------
def _plain(x):
    """Strip any leftover markup so old fragments render as text, not tags."""
    return re.sub(r"<[^>]+>", "", str(x)).replace("&rarr;", "->").replace("&amp;", "&").strip()


def move_tag(pct):
    """Plain text now; it used to be an HTML pill inside a hand-built row."""
    if pct is None: return "baru"
    return f"{'+' if pct >= 0 else ''}{pct:.1f}%"


def rows_html(items):
    """
    items: (tenant_id, name, sub, right) -> DataFrame.

    The name is unchanged so the 9 call sites need no edit, but it no longer
    returns markup. Hand-built rows of HTML were the single largest source of
    the overlapping layout, because Streamlit never measured them.
    """
    return pd.DataFrame([{"Nama": n, "Keterangan": s, "Nilai": _plain(r)}
                         for _, n, s, r in items])


def gauge_html(pct, label):
    """Returns the pair (value, label); render_card turns it into a metric."""
    return ("gauge", max(0.0, min(float(pct), 1.0)), label)


def render_card(title, inner, foot=None, dark=False):
    """
    A bordered container whose body is chosen by what it was handed:
    a DataFrame becomes a table, a gauge tuple becomes a metric with a
    progress bar, anything else is rendered as text.
    """
    with st.container(border=True):
        st.markdown(f"###### {title}")
        if isinstance(inner, pd.DataFrame):
            if not inner.empty:
                st.dataframe(inner, use_container_width=True, hide_index=True)
        elif isinstance(inner, tuple) and inner and inner[0] == "gauge":
            _, p, lb = inner
            st.metric("Capaian", f"{p*100:.0f}%")
            st.progress(p)
            st.caption(lb)
        elif inner:
            st.markdown(_plain(inner))
        if foot:
            st.caption(_plain(foot))



def render_header(subtitle=None):
    """No-op. The search bar and user chip were decoration built from raw HTML;
    the identity already lives in the sidebar, so the markup is simply gone."""
    return None


def render_page_head(title, sub=""):
    page_head(title, sub or None)


def initials(name):
    p = [x for x in str(name).replace("'", " ").split() if x]
    if not p: return "?"
    return (p[0][:2] if len(p) == 1 else p[0][0] + p[1][0]).upper()


def render_kpi(label, value, delta=None, variant="", caption=None, featured=False):
    """Compatibility shim: the old signature, a native metric underneath."""
    d = f"{delta:+.1f}%" if isinstance(delta, (int, float)) else None
    with st.container(border=True):
        st.metric(label, value, delta=d)
        if caption:
            st.caption(caption)


# ============================================================================
# AUTH PAGES
# ============================================================================
def get_user_email():
    """Try multiple methods to get the logged-in user's email."""
    try:
        if hasattr(st, "user") and hasattr(st.user, "email") and st.user.email:
            return st.user.email
    except: pass
    try:
        if hasattr(st, "user") and isinstance(st.user, dict) and st.user.get("email"):
            return st.user["email"]
    except: pass
    return st.session_state.get("manual_email", "")


def page_setup():
    """First-time setup — register the first Super Admin."""
    st.markdown('<div class="login-box"><h2>🌿 Initial Setup</h2><p>Buat akun Super Admin pertama.</p></div>', unsafe_allow_html=True)
    user_email = get_user_email()
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if user_email:
            st.info(f"Login sebagai: **{user_email}**")
        with st.form("setup_form"):
            if not user_email:
                user_email_input = st.text_input("Email Anda", placeholder="alfin@srkel.id")
            else:
                user_email_input = user_email
            display_name = st.text_input("Nama Lengkap", placeholder="e.g. Alfin")
            if st.form_submit_button("🚀 Buat Super Admin", use_container_width=True):
                email = user_email or user_email_input
                if not email or "@" not in email:
                    st.error("Email wajib diisi dengan format yang benar.")
                elif not display_name:
                    st.error("Nama lengkap wajib diisi.")
                else:
                    domain = email.split("@")[-1]
                    if domain not in ALLOWED_DOMAINS:
                        st.error(f"❌ Domain @{domain} tidak diizinkan.")
                    else:
                        get_db().create_user(email, display_name, "Super Admin", "ALL", "SYSTEM")
                        st.session_state["manual_email"] = email
                        st.success("✅ Super Admin berhasil dibuat!")
                        st.rerun()


def page_login():
    """Fallback login when Google auth doesn't provide email."""
    st.markdown('<div class="login-box"><h2>🌿 GROVE Sales Analytics</h2><p>Masukkan email terdaftar untuk melanjutkan.</p></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            email = st.text_input("Email", placeholder="nama@srkel.id")
            if st.form_submit_button("🔐 Login", use_container_width=True):
                if not email:
                    st.error("Email wajib diisi.")
                else:
                    domain = email.split("@")[-1] if "@" in email else ""
                    if domain not in ALLOWED_DOMAINS:
                        st.error(f"❌ Domain @{domain} tidak diizinkan.")
                    else:
                        db = get_db()
                        user = db.get_user(email)
                        if not user:
                            st.error("Email tidak terdaftar. Hubungi Super Admin.")
                        elif str(user.get("is_active","TRUE")).upper() not in ("TRUE","1"):
                            st.error("Akun dinonaktifkan. Hubungi Admin.")
                        else:
                            st.session_state["manual_email"] = email
                            st.session_state["authenticated"] = True
                            st.rerun()


def check_auth():
    """Check auth — Google email or fallback manual login. Returns user dict or None."""
    user_email = get_user_email()
    if not user_email:
        return None

    domain = user_email.split("@")[-1] if "@" in user_email else ""
    if domain not in ALLOWED_DOMAINS:
        st.error(f"❌ Email {user_email} — domain @{domain} tidak diizinkan.")
        st.stop()
        return None

    db = get_db()
    user = db.get_user(user_email)
    if not user:
        if not db.has_users():
            return None
        # Auto-create Viewer for Google-auth users only
        google_email = ""
        try:
            if hasattr(st, "user") and hasattr(st.user, "email") and st.user.email:
                google_email = st.user.email
        except: pass
        if google_email:
            name = user_email.split("@")[0].replace("."," ").title()
            db.create_user(user_email, name, "Viewer", "ALL", "AUTO")
            user = db.get_user(user_email)
        else:
            return None

    if str(user.get("is_active","TRUE")).upper() not in ("TRUE","1"):
        st.error("❌ Akun Anda dinonaktifkan. Hubungi Admin.")
        st.stop()
        return None

    return user


# ============================================================================
# SIDEBAR
# ============================================================================
def render_sidebar():
    user = st.session_state["user"]
    role = user["role"]
    with st.sidebar:
        stats = get_db().get_db_stats()
        st.markdown("### GROVE")
        st.caption(f"{user['display_name']} · {role}")
        st.caption(f"F&B {stats['sales_rows']:,} · Playground {stats['pg_rows']:,}")
        st.divider()

        menu = [
            "🏠 Master Dashboard",
            "🍽️ Dashboard F&B",
            "🎪 Dashboard Playground",
            "📈 Performa Tenant",
            "───────────────",
            "📤 Upload F&B (ESB)",
            "📤 Upload Playground",
        ]
        if role in ("Super Admin", "Admin"):
            menu.append("🏢 Kelola Tenant")
            menu.append("🏬 Kelola Unit & Sewa")
        if role == "Super Admin":
            menu.append("👥 Kelola User")
        menu.append("📋 Upload Log")

        selected = st.radio("Menu", menu, label_visibility="collapsed")
        if selected.startswith("──"):
            selected = "🏠 Master Dashboard"

        st.divider()
        if st.button("🚪 Logout", use_container_width=True):
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()
    return selected


# ============================================================================
# DASHBOARD HELPERS
# ============================================================================
def extract_hr(h):
    m = re.match(r"(\d{1,2}):", str(h))
    return int(m.group(1)) if m else 0

def assign_segment(hr):
    for seg, (s, e) in TIME_SEGMENTS.items():
        if s <= hr < e:
            return seg
    return "Other"

def assign_week_number(d, base_date=None):
    return (d.day - 1) // 7 + 1

def is_weekend(d):
    return d.weekday() >= 5

def enrich_df(df):
    df = df.copy()
    df["hr"] = df["sales_hour"].apply(extract_hr)
    df["segment"] = df["hr"].apply(assign_segment)
    df["is_weekend"] = df["sales_date"].apply(is_weekend)
    df["day_type"] = df["is_weekend"].map({True: "Weekend", False: "Weekday"})
    df["weekday"] = df["sales_date"].dt.day_name()
    df["week_num"] = df["sales_date"].apply(assign_week_number)
    df["week_label"] = "Week " + df["week_num"].astype(str)
    df["month"] = df["sales_date"].dt.to_period("M").astype(str)
    df["date_only"] = df["sales_date"].dt.date
    return df

PLT = dict()


PLT_FONT = dict(font=dict(family="ui-rounded, Segoe UI, system-ui, sans-serif", size=12))


# ============================================================================
# DASHBOARD EXPORT
# ============================================================================
def generate_dashboard_xlsx(df, sel_tenant, date_start, date_end):
    """Generate professional multi-sheet XLSX with openpyxl native charts."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, BarChart3D, LineChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint

    wb = Workbook()

    GD = "1B4332"; GM = "2D6A4F"; GL = "D8F3DC"; GLD = "D4A843"
    hf = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor=GD)
    bf = Font(name="Arial", size=10, color="333333")
    bfb = Font(name="Arial", size=10, bold=True, color="333333")
    tf = Font(name="Arial", size=14, bold=True, color=GD)
    sf = Font(name="Arial", size=10, color="888888", italic=True)
    rp_fmt = "#,##0"
    bdr = Border(left=Side("thin",color="DDDDDD"), right=Side("thin",color="DDDDDD"),
                 top=Side("thin",color="DDDDDD"), bottom=Side("thin",color="DDDDDD"))
    wrap_c = Alignment(wrap_text=True, vertical="center")
    cx = Alignment(horizontal="center", vertical="center")
    gray_fill = PatternFill("solid", fgColor="F5F5F5")

    chart_colors = ["2D6A4F","52B788","D4A843","457B9D","E76F51","E63946","6A4C93","1B4332"]

    def add_header(ws, title, subtitle, cols):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
        ws.cell(1, 1, title).font = tf; ws.row_dimensions[1].height = 28
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
        ws.cell(2, 1, subtitle).font = sf

    def write_table(ws, start_row, headers, data_rows, num_cols=None):
        r = start_row
        for c, h in enumerate(headers, 1):
            cell = ws.cell(r, c, h)
            cell.font = hf; cell.fill = hfill; cell.alignment = cx; cell.border = bdr
        ws.row_dimensions[r].height = 24
        for i, row_data in enumerate(data_rows):
            r = start_row + 1 + i
            for c, val in enumerate(row_data, 1):
                cell = ws.cell(r, c, val)
                cell.font = bf; cell.border = bdr; cell.alignment = wrap_c
                if num_cols and c in num_cols: cell.number_format = rp_fmt
                if i % 2 == 0: cell.fill = gray_fill
        return start_row + 1 + len(data_rows)

    def make_bar(ws, title, data_col, cat_col, min_row, max_row, anchor,
                 width=22, height=12, colors=None, x_title="", y_title="Nilai (Rp)", show_legend=False):
        c = BarChart(); c.type = "col"; c.grouping = "clustered"; c.style = 10
        c.title = title; c.width = width; c.height = height
        if y_title: c.y_axis.title = y_title
        if x_title: c.x_axis.title = x_title
        data = Reference(ws, min_col=data_col, min_row=min_row, max_row=max_row)
        cats = Reference(ws, min_col=cat_col, min_row=min_row+1, max_row=max_row)
        c.add_data(data, titles_from_data=True); c.set_categories(cats)
        if not show_legend: c.legend = None
        if colors:
            for idx, clr in enumerate(colors):
                if idx < len(c.series): c.series[idx].graphicalProperties.solidFill = clr
        elif c.series:
            c.series[0].graphicalProperties.solidFill = GM
        c.y_axis.numFmt = "#,##0"
        # Data labels
        for s in c.series:
            s.dLbls = DataLabelList()
            s.dLbls.showVal = True; s.dLbls.numFmt = "#,##0"
        ws.add_chart(c, anchor)
        return c

    def make_grouped_bar(ws, title, data_cols, cat_col, min_row, max_row, anchor,
                         width=24, height=13, colors=None, x_title="", y_title="Nilai"):
        c = BarChart(); c.type = "col"; c.grouping = "clustered"; c.style = 10
        c.title = title; c.width = width; c.height = height
        if y_title: c.y_axis.title = y_title
        if x_title: c.x_axis.title = x_title
        cats = Reference(ws, min_col=cat_col, min_row=min_row+1, max_row=max_row)
        for dc in data_cols:
            data = Reference(ws, min_col=dc, min_row=min_row, max_row=max_row)
            c.add_data(data, titles_from_data=True)
        c.set_categories(cats)
        if colors:
            for idx, clr in enumerate(colors):
                if idx < len(c.series): c.series[idx].graphicalProperties.solidFill = clr
        c.y_axis.numFmt = "#,##0"
        # Data labels
        for s in c.series:
            s.dLbls = DataLabelList()
            s.dLbls.showVal = True; s.dLbls.numFmt = "#,##0"
        ws.add_chart(c, anchor)
        return c

    def make_line(ws, title, data_cols, cat_col, min_row, max_row, anchor,
                  width=24, height=13, colors=None, x_title="Tanggal", y_title="Nilai (Rp)"):
        c = LineChart(); c.style = 10; c.title = title; c.width = width; c.height = height
        if y_title: c.y_axis.title = y_title
        if x_title: c.x_axis.title = x_title
        cats = Reference(ws, min_col=cat_col, min_row=min_row+1, max_row=max_row)
        for dc in data_cols:
            data = Reference(ws, min_col=dc, min_row=min_row, max_row=max_row)
            c.add_data(data, titles_from_data=True)
        c.set_categories(cats)
        if colors:
            for idx, clr in enumerate(colors):
                if idx < len(c.series):
                    c.series[idx].graphicalProperties.line.solidFill = clr
                    c.series[idx].graphicalProperties.line.width = 25000
        c.y_axis.numFmt = "#,##0"
        ws.add_chart(c, anchor)
        return c

    def make_pie(ws, title, data_col, cat_col, min_row, max_row, anchor, width=16, height=12):
        c = PieChart(); c.title = title; c.style = 10; c.width = width; c.height = height
        data = Reference(ws, min_col=data_col, min_row=min_row, max_row=max_row)
        cats = Reference(ws, min_col=cat_col, min_row=min_row+1, max_row=max_row)
        c.add_data(data, titles_from_data=True); c.set_categories(cats)
        c.dataLabels = DataLabelList()
        c.dataLabels.showPercent = True; c.dataLabels.showVal = True
        c.dataLabels.showCatName = True; c.dataLabels.numFmt = "#,##0"
        for i, clr in enumerate(chart_colors):
            if i < max_row - min_row:
                pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = clr; c.series[0].data_points.append(pt)
        ws.add_chart(c, anchor)

    tenant_label = sel_tenant or "All Tenants"
    period_label = f"{date_start} s.d. {date_end}"

    # ======== SHEET 1: EXECUTIVE SUMMARY ========
    ws1 = wb.active; ws1.title = "Executive Summary"; ws1.sheet_properties.tabColor = GD
    for c, w in [(1,30),(2,22)]: ws1.column_dimensions[chr(64+c)].width = w
    add_header(ws1, "GROVE Sales Analytics — Executive Summary", f"Tenant: {tenant_label}  |  Periode: {period_label}", 2)

    t_nett=df["nett_sales"].sum(); t_pax=df["pax_total"].sum()
    t_sub=df["subtotal"].sum(); t_disc=df["discount_total"].sum()
    n_days=df["date_only"].nunique()
    kpis = [("Total Nett Sales",t_nett,rp_fmt),("Total Pax",t_pax,"#,##0"),
            ("Average per Pax",t_nett/t_pax if t_pax else 0,rp_fmt),
            ("Total Subtotal",t_sub,rp_fmt),("Total Discount",t_disc,rp_fmt),
            ("Discount Rate",t_disc/t_sub if t_sub else 0,"0.0%"),
            ("Hari Aktif",n_days,"#,##0"),("Avg Daily Sales",t_nett/n_days if n_days else 0,rp_fmt),
            ("Jumlah Tenant",df["tenant_name"].nunique(),"#,##0")]
    r = 4
    ws1.cell(r,1,"Metrik").font=hf; ws1.cell(r,1).fill=hfill; ws1.cell(r,1).border=bdr
    ws1.cell(r,2,"Nilai").font=hf; ws1.cell(r,2).fill=hfill; ws1.cell(r,2).border=bdr
    for i,(lab,val,fmt) in enumerate(kpis):
        r=5+i; ws1.cell(r,1,lab).font=bfb; ws1.cell(r,1).border=bdr
        ws1.cell(r,2,val).font=Font(name="Arial",size=12,bold=True,color=GD)
        ws1.cell(r,2).number_format=fmt; ws1.cell(r,2).border=bdr
        if i%2==0: ws1.cell(r,1).fill=gray_fill; ws1.cell(r,2).fill=gray_fill

    # ======== SHEET 2: TREND HARIAN ========
    ws2 = wb.create_sheet("Trend Harian"); ws2.sheet_properties.tabColor = GM
    for c,w in [(1,14),(2,8),(3,10),(4,15),(5,15),(6,15)]: ws2.column_dimensions[chr(64+c)].width = w
    add_header(ws2, "Trend Harian", f"Tenant: {tenant_label}  |  Periode: {period_label}", 6)
    daily = df.groupby("date_only").agg(pax=("pax_total","sum"),sub=("subtotal","sum"),
        disc=("discount_total","sum"),nett=("nett_sales","sum")).reset_index()
    daily["wd"] = pd.to_datetime(daily["date_only"]).dt.strftime("%a")
    d_rows = [[str(r2["date_only"]),r2["wd"],r2["pax"],r2["sub"],r2["disc"],r2["nett"]] for _,r2 in daily.iterrows()]
    end_r = write_table(ws2, 4, ["Tanggal","Hari","Pax","Subtotal","Discount","Nett Sales"], d_rows, num_cols={3,4,5,6})

    # Bar: Nett Sales + Line: Pax
    make_bar(ws2, "Daily Nett Sales", 6, 1, 4, end_r-1, f"A{end_r+1}", x_title="Tanggal", y_title="Nett Sales (Rp)")
    make_line(ws2, "Daily Pax Trend", [3], 1, 4, end_r-1, f"A{end_r+16}", colors=[GLD], x_title="Tanggal", y_title="Jumlah Pax")

    # Weekday average
    wd_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    wd_avg = df.groupby("weekday")["nett_sales"].mean().reindex(wd_order).fillna(0).reset_index()
    wd_avg.columns = ["Day","Avg Sales"]
    r_wd = end_r + 31
    ws2.cell(r_wd, 1, "Rata-rata Sales per Hari").font = Font(name="Arial",size=12,bold=True,color=GD)
    r_wd += 1
    wd_rows = [[row["Day"], round(row["Avg Sales"])] for _,row in wd_avg.iterrows()]
    end_wd = write_table(ws2, r_wd, ["Hari","Avg Sales (Rp)"], wd_rows, num_cols={2})
    make_bar(ws2, "Rata-rata Sales per Hari", 2, 1, r_wd, end_wd-1, f"A{end_wd+1}", x_title="Hari", y_title="Avg Sales (Rp)")

    # ======== SHEET 3: ANALISIS PER JAM ========
    ws3 = wb.create_sheet("Analisis Per Jam"); ws3.sheet_properties.tabColor = "457B9D"
    for c,w in [(1,10),(2,12),(3,12),(4,15),(5,15)]: ws3.column_dimensions[chr(64+c)].width = w
    add_header(ws3, "Analisis Per Jam — Weekday vs Weekend", f"Tenant: {tenant_label}  |  Periode: {period_label}", 5)

    # WD vs WE pivot for charting
    hr_piv = df.groupby(["hr","day_type"])["pax_total"].sum().reset_index()
    hrs = sorted(hr_piv["hr"].unique())
    r_h = 4
    ws3.cell(r_h,1,"Jam").font=hf; ws3.cell(r_h,1).fill=hfill; ws3.cell(r_h,1).border=bdr
    ws3.cell(r_h,2,"Weekday Pax").font=hf; ws3.cell(r_h,2).fill=hfill; ws3.cell(r_h,2).border=bdr
    ws3.cell(r_h,3,"Weekend Pax").font=hf; ws3.cell(r_h,3).fill=hfill; ws3.cell(r_h,3).border=bdr
    ws3.cell(r_h,4,"Weekday Sales").font=hf; ws3.cell(r_h,4).fill=hfill; ws3.cell(r_h,4).border=bdr
    ws3.cell(r_h,5,"Weekend Sales").font=hf; ws3.cell(r_h,5).fill=hfill; ws3.cell(r_h,5).border=bdr
    hr_sales = df.groupby(["hr","day_type"])["nett_sales"].sum().reset_index()
    for i,h in enumerate(hrs):
        r = 5 + i
        ws3.cell(r,1,f"{h:02d}:00").font=bf; ws3.cell(r,1).border=bdr
        wd_p = hr_piv[(hr_piv["hr"]==h)&(hr_piv["day_type"]=="Weekday")]["pax_total"].sum()
        we_p = hr_piv[(hr_piv["hr"]==h)&(hr_piv["day_type"]=="Weekend")]["pax_total"].sum()
        wd_s = hr_sales[(hr_sales["hr"]==h)&(hr_sales["day_type"]=="Weekday")]["nett_sales"].sum()
        we_s = hr_sales[(hr_sales["hr"]==h)&(hr_sales["day_type"]=="Weekend")]["nett_sales"].sum()
        for c,v in [(2,wd_p),(3,we_p),(4,wd_s),(5,we_s)]:
            ws3.cell(r,c,v).font=bf; ws3.cell(r,c).border=bdr; ws3.cell(r,c).number_format=rp_fmt
        if i%2==0:
            for c in range(1,6): ws3.cell(r,c).fill=gray_fill
    end_h = 5 + len(hrs)
    make_grouped_bar(ws3, "Pax per Jam — Weekday vs Weekend", [2,3], 1, 4, end_h-1, f"A{end_h+1}",
                     colors=["457B9D","E63946"], x_title="Jam", y_title="Jumlah Pax")
    make_grouped_bar(ws3, "Sales per Jam — Weekday vs Weekend", [4,5], 1, 4, end_h-1, f"A{end_h+16}",
                     colors=["457B9D","E63946"], x_title="Jam", y_title="Sales (Rp)")

    # Traffic summary
    r_tr = end_h + 31
    ws3.cell(r_tr,1,"Ringkasan Traffic").font=Font(name="Arial",size=12,bold=True,color=GD)
    r_tr += 1
    trf = df.groupby("day_type").agg(pax=("pax_total","sum"),sales=("nett_sales","sum")).reset_index()
    tr_rows = [[row["day_type"],row["pax"],row["sales"]] for _,row in trf.iterrows()]
    write_table(ws3, r_tr, ["Day Type","Total Pax","Total Sales (Rp)"], tr_rows, num_cols={2,3})

    # ======== SHEET 4: PERBANDINGAN TENANT ========
    ws4 = wb.create_sheet("Perbandingan Tenant"); ws4.sheet_properties.tabColor = GLD
    for c,w in [(1,25),(2,10),(3,12),(4,15),(5,15),(6,12),(7,10)]: ws4.column_dimensions[chr(64+c)].width = w
    add_header(ws4, "Perbandingan Tenant", f"Periode: {period_label}  |  Total: Rp {t_nett:,.0f}", 7)
    ta = df.groupby("tenant_name").agg(days=("date_only","nunique"),pax=("pax_total","sum"),
        sub=("subtotal","sum"),disc=("discount_total","sum"),nett=("nett_sales","sum")).reset_index()
    ta["avg_pax"]=(ta["nett"]/ta["pax"].replace(0,1)).fillna(0).round(0); ta["disc_r"]=(ta["disc"]/ta["sub"].replace(0,1)).fillna(0)
    ta = ta.sort_values("nett",ascending=False)
    ta_rows = [[r2["tenant_name"],r2["days"],r2["pax"],r2["sub"],r2["nett"],r2["avg_pax"],r2["disc_r"]] for _,r2 in ta.iterrows()]
    end_r = write_table(ws4, 4, ["Tenant","Hari","Pax","Subtotal","Nett Sales","Avg/Pax","Disc %"], ta_rows, num_cols={3,4,5,6})
    for r in range(5,5+len(ta_rows)): ws4.cell(r,7).number_format="0.0%"
    make_bar(ws4, "Total Sales per Tenant", 5, 1, 4, end_r-1, f"A{end_r+1}",
             colors=[chart_colors[i%len(chart_colors)] for i in range(len(ta_rows))], x_title="Tenant", y_title="Nett Sales (Rp)")
    if len(ta_rows) > 1:
        make_pie(ws4, "Kontribusi Sales per Tenant", 5, 1, 4, end_r-1, f"A{end_r+16}")

    # ======== SHEET 5: TIME SEGMENT ========
    ws5 = wb.create_sheet("Time Segment"); ws5.sheet_properties.tabColor = "E76F51"
    for c,w in [(1,25),(2,15),(3,12),(4,12),(5,15),(6,12)]: ws5.column_dimensions[chr(64+c)].width = w
    add_header(ws5, "Time Segment Analysis", f"Tenant: {tenant_label}  |  Periode: {period_label}", 6)
    seg_df2 = df[df["segment"]!="Other"]
    # Segment summary pivot
    seg_sum = seg_df2.groupby(["segment","day_type"]).agg(pax=("pax_total","sum"),sales=("nett_sales","sum")).reset_index()
    seg_order = ["Breakfast","Lunch","After Office"]
    r_s = 4
    ws5.cell(r_s,1,"Segment").font=hf;ws5.cell(r_s,1).fill=hfill;ws5.cell(r_s,1).border=bdr
    ws5.cell(r_s,2,"Weekday Pax").font=hf;ws5.cell(r_s,2).fill=hfill;ws5.cell(r_s,2).border=bdr
    ws5.cell(r_s,3,"Weekend Pax").font=hf;ws5.cell(r_s,3).fill=hfill;ws5.cell(r_s,3).border=bdr
    ws5.cell(r_s,4,"Weekday Sales").font=hf;ws5.cell(r_s,4).fill=hfill;ws5.cell(r_s,4).border=bdr
    ws5.cell(r_s,5,"Weekend Sales").font=hf;ws5.cell(r_s,5).fill=hfill;ws5.cell(r_s,5).border=bdr
    for i,seg in enumerate(seg_order):
        r = 5 + i
        ws5.cell(r,1,seg).font=bfb;ws5.cell(r,1).border=bdr
        wd_p = seg_sum[(seg_sum["segment"]==seg)&(seg_sum["day_type"]=="Weekday")]["pax"].sum()
        we_p = seg_sum[(seg_sum["segment"]==seg)&(seg_sum["day_type"]=="Weekend")]["pax"].sum()
        wd_s = seg_sum[(seg_sum["segment"]==seg)&(seg_sum["day_type"]=="Weekday")]["sales"].sum()
        we_s = seg_sum[(seg_sum["segment"]==seg)&(seg_sum["day_type"]=="Weekend")]["sales"].sum()
        for c,v in [(2,wd_p),(3,we_p),(4,wd_s),(5,we_s)]:
            ws5.cell(r,c,v).font=bf;ws5.cell(r,c).border=bdr;ws5.cell(r,c).number_format=rp_fmt
        if i%2==0:
            for c in range(1,6): ws5.cell(r,c).fill=gray_fill
    end_s = 5 + len(seg_order)
    make_grouped_bar(ws5, "Pax by Segment — Weekday vs Weekend", [2,3], 1, 4, end_s-1, f"A{end_s+1}", colors=["457B9D","E63946"], x_title="Segment", y_title="Jumlah Pax")
    make_grouped_bar(ws5, "Sales by Segment — Weekday vs Weekend", [4,5], 1, 4, end_s-1, f"A{end_s+16}", colors=["457B9D","E63946"], x_title="Segment", y_title="Sales (Rp)")

    # Detail per tenant per segment
    r_det = end_s + 31
    ws5.cell(r_det,1,"Detail per Tenant per Segment").font=Font(name="Arial",size=12,bold=True,color=GD)
    ws5.merge_cells(start_row=r_det,start_column=1,end_row=r_det,end_column=6)
    r_det += 1
    seg_ta = seg_df2.groupby(["tenant_name","segment"]).agg(pax=("pax_total","sum"),sales=("nett_sales","sum")).reset_index()
    st_rows = [[r2["tenant_name"],r2["segment"],r2["pax"],r2["sales"],
                r2["sales"]/r2["pax"] if r2["pax"] else 0] for _,r2 in seg_ta.sort_values(["segment","tenant_name"]).iterrows()]
    write_table(ws5, r_det, ["Tenant","Segment","Pax","Sales (Rp)","Avg/Pax"], st_rows, num_cols={3,4,5})

    # ======== SHEET 6: WEEKLY REPORT ========
    ws6 = wb.create_sheet("Weekly Report"); ws6.sheet_properties.tabColor = "6A4C93"
    for c,w in [(1,25),(2,10),(3,12),(4,15),(5,15),(6,15)]: ws6.column_dimensions[chr(64+c)].width = w
    add_header(ws6, "Weekly Report", f"Tenant: {tenant_label}  |  Periode: {period_label}", 6)
    wk = df.groupby(["tenant_name","week_label"]).agg(pax=("pax_total","sum"),sub=("subtotal","sum"),
        disc=("discount_total","sum"),nett=("nett_sales","sum")).reset_index()
    w_rows = [[r2["tenant_name"],r2["week_label"],r2["pax"],r2["sub"],r2["disc"],r2["nett"]]
              for _,r2 in wk.sort_values(["week_label","tenant_name"]).iterrows()]
    end_r = write_table(ws6, 4, ["Tenant","Week","Pax","Subtotal","Discount","Nett Sales"], w_rows, num_cols={3,4,5,6})

    # Weekly pivot for chart
    weeks = sorted(df["week_label"].unique())
    tenants = sorted(df["tenant_name"].unique())
    r_wp = end_r + 1
    ws6.cell(r_wp,1,"Pivot: Sales per Tenant per Week").font=Font(name="Arial",size=12,bold=True,color=GD)
    r_wp += 1
    ws6.cell(r_wp,1,"Tenant").font=hf;ws6.cell(r_wp,1).fill=hfill;ws6.cell(r_wp,1).border=bdr
    for wi,wl in enumerate(weeks):
        ws6.cell(r_wp,2+wi,wl).font=hf;ws6.cell(r_wp,2+wi).fill=hfill;ws6.cell(r_wp,2+wi).border=bdr
    for ti,tn in enumerate(tenants):
        r = r_wp+1+ti
        ws6.cell(r,1,tn).font=bfb;ws6.cell(r,1).border=bdr
        for wi,wl in enumerate(weeks):
            val = wk[(wk["tenant_name"]==tn)&(wk["week_label"]==wl)]["nett"].sum()
            ws6.cell(r,2+wi,val).font=bf;ws6.cell(r,2+wi).border=bdr;ws6.cell(r,2+wi).number_format=rp_fmt
    end_wp = r_wp+1+len(tenants)
    if weeks:
        make_grouped_bar(ws6, "Tenant Sales per Week", list(range(2,2+len(weeks))), 1, r_wp, end_wp-1,
                         f"A{end_wp+1}", colors=["8B0000","CC3333","DAA520","B8860B","457B9D","6A4C93"],
                         x_title="Tenant", y_title="Nett Sales (Rp)")

    # Daily detail
    day_map = {"Monday":"Sen","Tuesday":"Sel","Wednesday":"Rab","Thursday":"Kam",
               "Friday":"Jum","Saturday":"Sab","Sunday":"Min"}
    r_dd = end_wp + 17
    ws6.cell(r_dd,1,"Detail Harian per Tenant").font=Font(name="Arial",size=12,bold=True,color=GD)
    ws6.merge_cells(start_row=r_dd,start_column=1,end_row=r_dd,end_column=6)
    r_dd += 1
    dd = df.groupby(["tenant_name","week_label","weekday"]).agg(pax=("pax_total","sum"),sales=("nett_sales","sum")).reset_index()
    dd_rows = [[r2["tenant_name"],r2["week_label"],day_map.get(r2["weekday"],r2["weekday"]),r2["pax"],r2["sales"]]
               for _,r2 in dd.sort_values(["tenant_name","week_label"]).iterrows()]
    write_table(ws6, r_dd, ["Tenant","Week","Hari","Pax","Sales (Rp)"], dd_rows, num_cols={4,5})

    # ======== SHEET 7: MONTHLY OVERVIEW ========
    ws7 = wb.create_sheet("Monthly Overview"); ws7.sheet_properties.tabColor = "E63946"
    for c,w in [(1,12),(2,25),(3,12),(4,15),(5,15),(6,15)]: ws7.column_dimensions[chr(64+c)].width = w
    add_header(ws7, "Monthly Overview", f"Tenant: {tenant_label}  |  Periode: {period_label}", 6)
    mt = df.groupby(["month","tenant_name"]).agg(pax=("pax_total","sum"),sub=("subtotal","sum"),
        disc=("discount_total","sum"),nett=("nett_sales","sum")).reset_index()
    m_rows = [[r2["month"],r2["tenant_name"],r2["pax"],r2["sub"],r2["disc"],r2["nett"]]
              for _,r2 in mt.sort_values(["month","tenant_name"]).iterrows()]
    end_r = write_table(ws7, 4, ["Bulan","Tenant","Pax","Subtotal","Discount","Nett Sales"], m_rows, num_cols={3,4,5,6})

    # Monthly grand total
    r_gt = end_r + 1
    ws7.cell(r_gt,1,"Grand Total per Bulan").font=Font(name="Arial",size=12,bold=True,color=GD)
    r_gt += 1
    mt_tot = df.groupby("month").agg(pax=("pax_total","sum"),nett=("nett_sales","sum")).reset_index().sort_values("month")
    gt_rows = [[r2["month"],r2["pax"],r2["nett"]] for _,r2 in mt_tot.iterrows()]
    end_gt = write_table(ws7, r_gt, ["Bulan","Total Pax","Nett Sales (Rp)"], gt_rows, num_cols={2,3})
    make_bar(ws7, "Total Nett Sales per Bulan", 3, 1, r_gt, end_gt-1, f"A{end_gt+1}", colors=["8B0000"]*12, x_title="Bulan", y_title="Nett Sales (Rp)")
    make_bar(ws7, "Total Pax per Bulan", 2, 1, r_gt, end_gt-1, f"A{end_gt+16}", colors=[GM]*12, x_title="Bulan", y_title="Jumlah Pax")

    # ======== SHEET 8: DEEP DIVE ========
    ws8 = wb.create_sheet("Deep Dive"); ws8.sheet_properties.tabColor = GD
    for c,w in [(1,14),(2,15),(3,15),(4,15)]: ws8.column_dimensions[chr(64+c)].width = w
    add_header(ws8, "Deep Dive — Moving Average & Statistics", f"Tenant: {tenant_label}  |  Periode: {period_label}", 4)
    dma = df.groupby("date_only")["nett_sales"].sum().reset_index()
    dma.columns = ["Tanggal","Sales"]; dma = dma.sort_values("Tanggal")
    dma["MA_7"] = dma["Sales"].rolling(7,min_periods=1).mean().round(0)
    dma["MA_14"] = dma["Sales"].rolling(14,min_periods=1).mean().round(0)
    ma_rows = [[str(r2["Tanggal"]),r2["Sales"],r2["MA_7"],r2["MA_14"]] for _,r2 in dma.iterrows()]
    end_r = write_table(ws8, 4, ["Tanggal","Daily Sales","MA 7-Day","MA 14-Day"], ma_rows, num_cols={2,3,4})
    make_line(ws8, "Moving Average Analysis", [2,3,4], 1, 4, end_r-1, f"A{end_r+1}",
              colors=["D8F3DC","52B788","1B4332"], x_title="Tanggal", y_title="Sales (Rp)")

    # Discount analysis
    r_disc = end_r + 17
    ws8.cell(r_disc,1,"Discount Analysis per Hari").font=Font(name="Arial",size=12,bold=True,color=GD)
    r_disc += 1
    dd2 = df.groupby("date_only").agg(disc=("discount_total","sum"),sales=("nett_sales","sum"),
        pax=("pax_total","sum")).reset_index()
    disc_rows = [[str(r2["date_only"]),r2["disc"],r2["sales"],r2["pax"]] for _,r2 in dd2.iterrows()]
    write_table(ws8, r_disc, ["Tanggal","Discount (Rp)","Sales (Rp)","Pax"], disc_rows, num_cols={2,3,4})

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def generate_dashboard_html(df, sel_tenant, date_start, date_end):
    """Generate self-contained HTML report with all Plotly charts — print to PDF via browser."""

    tenant_label = sel_tenant or "All Tenants"
    period_label = f"{date_start} s.d. {date_end}"
    t_nett = df["nett_sales"].sum(); t_pax = df["pax_total"].sum()
    t_sub = df["subtotal"].sum(); t_disc = df["discount_total"].sum()
    n_days = df["date_only"].nunique()
    avg_pax = t_nett / t_pax if t_pax else 0
    avg_daily = t_nett / n_days if n_days else 0
    disc_pct = t_disc / t_sub * 100 if t_sub else 0

    chart_h = 500
    chart_layout = dict(
        font=dict(family="ui-rounded, Segoe UI, system-ui, sans-serif",
                  size=13, color="#4A524D"),
        title_font=dict(size=16, color="#131715"),
        margin=dict(t=44, b=44, l=58, r=24), height=chart_h)

    charts_html = []

    def add_chart(fig, section_title=None):
        if section_title:
            charts_html.append(f'<div class="section-title">{section_title}</div>')
        charts_html.append(fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False}))
        charts_html.append('<div class="page-break"></div>')

    # --- KPI ---
    kpi_html = f"""
    <div class="kpi-row">
        <div class="kpi"><div class="kpi-label">Total Nett Sales</div><div class="kpi-value">Rp {t_nett:,.0f}</div></div>
        <div class="kpi gold"><div class="kpi-label">Total Pax</div><div class="kpi-value">{t_pax:,}</div></div>
        <div class="kpi blue"><div class="kpi-label">Avg / Pax</div><div class="kpi-value">Rp {avg_pax:,.0f}</div></div>
        <div class="kpi red"><div class="kpi-label">Discount Rate</div><div class="kpi-value">{disc_pct:.1f}%</div></div>
        <div class="kpi orange"><div class="kpi-label">Avg Daily Sales</div><div class="kpi-value">Rp {avg_daily:,.0f}</div></div>
    </div>
    """
    charts_html.append(kpi_html)

    # --- 1. TREND HARIAN ---
    daily = df.groupby("date_only").agg(nett=("nett_sales","sum"), pax=("pax_total","sum")).reset_index()
    fig1 = make_subplots(specs=[[{"secondary_y":True}]])
    fig1.add_trace(go.Bar(x=daily["date_only"].astype(str), y=daily["nett"], name="Nett Sales",
                          marker_color="#1A6B3F", opacity=0.85), secondary_y=False)
    fig1.add_trace(go.Scatter(x=daily["date_only"].astype(str), y=daily["pax"], name="Pax",
                              mode="lines+markers", line=dict(color="#C08A2C",width=2.5)), secondary_y=True)
    fig1.update_layout(title="Daily Sales & Pax Trend", hovermode="x unified", **chart_layout)
    fig1.update_yaxes(title_text="Nett Sales (Rp)", secondary_y=False)
    fig1.update_yaxes(title_text="Pax", secondary_y=True)
    add_chart(fig1, "1. Trend Harian")

    wd_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    wd = df.groupby("weekday")["nett_sales"].mean().reindex(wd_order).fillna(0).reset_index()
    wd.columns = ["Day","Avg"]
    fig_wd = px.bar(wd, x="Day", y="Avg", title="Rata-rata Sales per Hari",
                    color="Avg", color_continuous_scale=["#DCEEE4","#103D28"],
                    labels={"Day":"Hari","Avg":"Avg Sales (Rp)"})
    fig_wd.update_layout(showlegend=False, coloraxis_showscale=False, **chart_layout)
    add_chart(fig_wd)

    # --- 2. ANALISIS PER JAM ---
    hr_wd = df.groupby(["hr","day_type"])["pax_total"].sum().reset_index()
    hr_wd["lbl"] = hr_wd["hr"].apply(lambda x: f"{x:02d}:00")
    fig_h = px.bar(hr_wd, x="lbl", y="pax_total", color="day_type", barmode="group",
                   title="Transaction per Hour — Weekday vs Weekend",
                   color_discrete_map={"Weekday":"#3B4C7A","Weekend":"#C0483C"},
                   text="pax_total", labels={"lbl":"Jam","pax_total":"Pax","day_type":""})
    fig_h.update_traces(textposition="outside", textfont_size=11)
    fig_h.update_layout(**chart_layout)
    add_chart(fig_h, "2. Analisis Per Jam")

    traffic = df.groupby("day_type").agg(pax=("pax_total","sum"), sales=("nett_sales","sum")).reset_index()
    fig_trf = px.bar(traffic, x="day_type", y=["pax","sales"], barmode="group",
                     title="Tenant Traffic & Sales — Weekday vs Weekend",
                     labels={"day_type":"","value":"","variable":""})
    fig_trf.update_layout(**chart_layout)
    add_chart(fig_trf)

    hm = df.groupby(["weekday","hr"])["nett_sales"].sum().reset_index()
    hm_piv = hm.pivot(index="weekday", columns="hr", values="nett_sales").fillna(0).reindex(wd_order)
    fig_hm = px.imshow(hm_piv, aspect="auto", title="Heatmap: Hari × Jam",
                       color_continuous_scale=["#DCEEE4","#103D28"],
                       labels=dict(x="Jam",y="Hari",color="Sales (Rp)"))
    fig_hm.update_layout(**chart_layout)
    add_chart(fig_hm)

    # --- 3. PERBANDINGAN TENANT ---
    if df["tenant_name"].nunique() > 1:
        ta = df.groupby("tenant_name")["nett_sales"].sum().reset_index().sort_values("nett_sales",ascending=False)
        fig_tb = px.bar(ta, x="tenant_name", y="nett_sales", title="Total Sales per Tenant",
                        color="tenant_name", color_discrete_sequence=CHART_PALETTE, text="nett_sales",
                        labels={"tenant_name":"Tenant","nett_sales":"Nett Sales (Rp)"})
        fig_tb.update_traces(textposition="outside", texttemplate="Rp%{text:,.0f}", textfont_size=12)
        fig_tb.update_layout(showlegend=False, **chart_layout)
        add_chart(fig_tb, "3. Perbandingan Tenant")

        fig_pie = px.pie(ta, values="nett_sales", names="tenant_name", title="Kontribusi Sales per Tenant",
                         color_discrete_sequence=CHART_PALETTE, hole=0.4)
        fig_pie.update_layout(**chart_layout)
        add_chart(fig_pie)

        td = df.groupby(["date_only","tenant_name"])["nett_sales"].sum().reset_index()
        fig_td = px.line(td, x="date_only", y="nett_sales", color="tenant_name",
                         title="Daily Trend per Tenant", color_discrete_sequence=CHART_PALETTE,
                         labels={"date_only":"Tanggal","nett_sales":"Nett Sales (Rp)","tenant_name":"Tenant"})
        fig_td.update_layout(**chart_layout)
        add_chart(fig_td)

    # --- 4. TIME SEGMENT ---
    seg_df2 = df[df["segment"]!="Other"]
    if not seg_df2.empty:
        seg_pax = seg_df2.groupby(["segment","day_type"])["pax_total"].sum().reset_index()
        seg_order = ["Breakfast","Lunch","After Office"]
        fig_seg = px.bar(seg_pax, x="segment", y="pax_total", color="day_type", barmode="group",
                         title="Pax by Time Segment — Weekday vs Weekend",
                         color_discrete_map={"Weekday":"#3B4C7A","Weekend":"#C0483C"},
                         text="pax_total", category_orders={"segment":seg_order},
                         labels={"segment":"Segment","pax_total":"Pax","day_type":""})
        fig_seg.update_traces(textposition="outside", texttemplate="%{text:,}", textfont_size=13)
        fig_seg.update_layout(**chart_layout)
        add_chart(fig_seg, "4. Time Segment")

        seg_sales = seg_df2.groupby(["segment","day_type"])["nett_sales"].sum().reset_index()
        fig_ss = px.bar(seg_sales, x="segment", y="nett_sales", color="day_type", barmode="group",
                        title="Sales by Time Segment — Weekday vs Weekend",
                        color_discrete_map={"Weekday":"#3B4C7A","Weekend":"#C0483C"},
                        text="nett_sales", category_orders={"segment":seg_order},
                        labels={"segment":"Segment","nett_sales":"Sales (Rp)","day_type":""})
        fig_ss.update_traces(textposition="outside", texttemplate="Rp%{text:,.0f}", textfont_size=12)
        fig_ss.update_layout(**chart_layout)
        add_chart(fig_ss)

    # --- 5. WEEKLY REPORT ---
    wk = df.groupby(["tenant_name","week_label"])["nett_sales"].sum().reset_index()
    fig_wk = px.bar(wk, x="tenant_name", y="nett_sales", color="week_label", barmode="group",
                    title="Tenant Sales per Week", text="nett_sales",
                    color_discrete_sequence=["#8A3730","#CC3333","#DAA520","#B8860B","#3B4C7A"],
                    labels={"tenant_name":"Tenant","nett_sales":"Sales (Rp)","week_label":""})
    fig_wk.update_traces(textposition="outside", texttemplate="Rp%{text:,.0f}", textfont_size=10)
    fig_wk.update_layout(**{**chart_layout, "height": 550})
    add_chart(fig_wk, "5. Weekly Report")

    # Per-tenant weekly detail
    day_map = {"Monday":"Sen","Tuesday":"Sel","Wednesday":"Rab","Thursday":"Kam",
               "Friday":"Jum","Saturday":"Sab","Sunday":"Min"}
    day_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    last_week = sorted(df["week_label"].unique())[-1] if df["week_label"].nunique() > 0 else ""
    if last_week:
        wdf = df[df["week_label"]==last_week]
        for tn in wdf["tenant_name"].unique():
            tdf = wdf[wdf["tenant_name"]==tn]
            ds = tdf.groupby("weekday")["nett_sales"].sum().reindex(day_order).fillna(0)
            avg_val = ds.mean()
            day_chart = pd.DataFrame({"Hari":[day_map.get(d,d) for d in day_order], "Sales":ds.values})
            fig_d = px.bar(day_chart, x="Hari", y="Sales", title=f"{tn} — {last_week} (Avg: Rp {avg_val:,.0f})",
                           text="Sales", color_discrete_sequence=["#6495ED"],
                           labels={"Hari":"Hari","Sales":"Sales (Rp)"})
            fig_d.update_traces(textposition="outside", texttemplate="Rp%{text:,.0f}", textfont_size=11)
            fig_d.update_layout(**{**chart_layout, "height": 400})
            add_chart(fig_d)

    # --- 6. MONTHLY OVERVIEW ---
    mt_tot = df.groupby("month").agg(pax=("pax_total","sum"), sales=("nett_sales","sum")).reset_index().sort_values("month")
    fig_mt = px.bar(mt_tot, x="month", y="sales", title="Total Nett Sales per Bulan",
                    text="sales", color_discrete_sequence=["#8A3730"],
                    labels={"month":"Bulan","sales":"Nett Sales (Rp)"})
    fig_mt.update_traces(textposition="outside", texttemplate="Rp%{text:,.0f}", textfont_size=13)
    fig_mt.update_layout(**chart_layout)
    add_chart(fig_mt, "6. Monthly Overview")

    fig_mp = px.bar(mt_tot, x="month", y="pax", title="Total Pax per Bulan",
                    text="pax", color_discrete_sequence=["#1A6B3F"],
                    labels={"month":"Bulan","pax":"Jumlah Pax"})
    fig_mp.update_traces(textposition="outside", texttemplate="%{text:,}", textfont_size=13)
    fig_mp.update_layout(**chart_layout)
    add_chart(fig_mp)

    # --- 7. DEEP DIVE ---
    dma = df.groupby("date_only")["nett_sales"].sum().reset_index()
    dma.columns = ["Tanggal","Sales"]; dma = dma.sort_values("Tanggal")
    dma["MA_7"] = dma["Sales"].rolling(7,min_periods=1).mean()
    dma["MA_14"] = dma["Sales"].rolling(14,min_periods=1).mean()
    fig_ma = go.Figure()
    fig_ma.add_trace(go.Scatter(x=dma["Tanggal"].astype(str), y=dma["Sales"], name="Daily",
                                mode="lines", line=dict(color="#DCEEE4",width=1)))
    fig_ma.add_trace(go.Scatter(x=dma["Tanggal"].astype(str), y=dma["MA_7"], name="MA 7-day",
                                mode="lines", line=dict(color="#2E9160",width=2.5)))
    fig_ma.add_trace(go.Scatter(x=dma["Tanggal"].astype(str), y=dma["MA_14"], name="MA 14-day",
                                mode="lines", line=dict(color="#103D28",width=2.5)))
    fig_ma.update_layout(title="Moving Average Analysis", xaxis_title="Tanggal", yaxis_title="Sales (Rp)", **chart_layout)
    add_chart(fig_ma, "7. Deep Dive")

    dd2 = df.groupby("date_only").agg(disc=("discount_total","sum"),sales=("nett_sales","sum"),pax=("pax_total","sum")).reset_index()
    fig_dsc = px.scatter(dd2, x="disc", y="sales", size="pax", title="Discount vs Sales Impact",
                         color="pax", color_continuous_scale=["#DCEEE4","#103D28"],
                         labels={"disc":"Discount (Rp)","sales":"Sales (Rp)","pax":"Pax"})
    fig_dsc.update_layout(**chart_layout)
    add_chart(fig_dsc)

    # --- SUMMARY TABLE ---
    sdf = df.groupby("tenant_name").agg(days=("date_only","nunique"),pax=("pax_total","sum"),
        sub=("subtotal","sum"),disc=("discount_total","sum"),nett=("nett_sales","sum")).reset_index()
    sdf["avg_pax"] = (sdf["nett"]/sdf["pax"].replace(0,1)).round(0)
    sdf["disc_rate"] = (sdf["disc"]/sdf["sub"].replace(0,1)*100).round(1)
    sdf["avg_daily"] = (sdf["nett"]/sdf["days"].replace(0,1)).round(0)
    table_html = '<div class="section-title">Ringkasan Statistik</div><table>'
    table_html += "<tr>" + "".join(f"<th>{h}</th>" for h in ["Tenant","Hari","Pax","Subtotal","Discount","Nett Sales","Avg/Pax","Disc %","Avg Daily"]) + "</tr>"
    for _,r in sdf.iterrows():
        table_html += f'<tr><td>{r["tenant_name"]}</td><td>{r["days"]}</td><td>{r["pax"]:,.0f}</td>'
        table_html += f'<td>Rp {r["sub"]:,.0f}</td><td>Rp {r["disc"]:,.0f}</td><td>Rp {r["nett"]:,.0f}</td>'
        table_html += f'<td>Rp {r["avg_pax"]:,.0f}</td><td>{r["disc_rate"]:.1f}%</td><td>Rp {r["avg_daily"]:,.0f}</td></tr>'
    table_html += "</table>"
    charts_html.append(table_html)

    # --- ASSEMBLE HTML ---
    body = "\\n".join(charts_html)
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>GROVE Sales Analytics — Dashboard Report</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');
* {{ font-family: ui-rounded, 'Segoe UI', system-ui, -apple-system, sans-serif; margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #fff; padding: 30px 40px; color: #333; max-width: 1100px; margin: 0 auto; }}
.report-header {{
    background: linear-gradient(135deg, #103D28, #1A6B3F, #2E9160);
    color: white; padding: 24px 32px; border-radius: 12px; margin-bottom: 24px;
}}
.report-header h1 {{ font-size: 28px; font-weight: 800; margin-bottom: 4px; }}
.report-header p {{ font-size: 14px; opacity: 0.85; }}
.kpi-row {{ display: flex; gap: 12px; margin-bottom: 24px; flex-wrap: wrap; }}
.kpi {{
    flex: 1; min-width: 160px; background: #fff; border-radius: 10px; padding: 16px;
    border-left: 5px solid #1A6B3F; box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}
.kpi.gold {{ border-left-color: #C08A2C; }}
.kpi.blue {{ border-left-color: #3B4C7A; }}
.kpi.red {{ border-left-color: #C0483C; }}
.kpi.orange {{ border-left-color: #9E6B4A; }}
.kpi-label {{ font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }}
.kpi-value {{ font-size: 22px; font-weight: 800; color: #103D28; margin-top: 4px; }}
.section-title {{
    font-size: 20px; font-weight: 700; color: #103D28;
    border-bottom: 3px solid #2E9160; padding-bottom: 6px;
    margin: 32px 0 16px;
}}
table {{ width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 13px; }}
th {{ background: #103D28; color: white; padding: 10px 12px; text-align: left; }}
td {{ padding: 8px 12px; border-bottom: 1px solid #eee; }}
tr:nth-child(even) {{ background: #f9f9f9; }}
.page-break {{ page-break-after: auto; }}
.print-btn {{
    position: fixed; top: 16px; right: 16px; z-index: 999;
    background: #1A6B3F; color: white; border: none; padding: 12px 24px;
    border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}}
.print-btn:hover {{ background: #103D28; }}
@media print {{
    .print-btn {{ display: none; }}
    body {{ padding: 10px; }}
    .report-header {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .kpi {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    th {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .page-break {{ page-break-after: always; }}
}}
</style></head><body>
<button class="print-btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
<div class="report-header">
    <h1>📊 GROVE Sales Analytics</h1>
    <p>Dashboard Report — Tenant: {tenant_label} | Periode: {period_label} | Generated: {datetime.now().strftime("%d %B %Y, %H:%M")}</p>
</div>
{body}
<div style="text-align:center; margin-top:40px; padding:16px; color:#888; font-size:12px; border-top:1px solid #eee;">
    GROVE Sales Analytics v{APP_VERSION} — Generated {datetime.now().strftime("%d/%m/%Y %H:%M")}
</div>
</body></html>"""

    return html.encode("utf-8")


def generate_playground_xlsx(df, date_start, date_end):
    """Generate Playground multi-sheet XLSX with charts."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint

    wb = Workbook()
    PG = "6A4C93"; PGL = "E8D5F5"; GLD = "D4A843"
    hf = Font(name="Arial",size=11,bold=True,color="FFFFFF")
    hfill = PatternFill("solid",fgColor=PG)
    bf = Font(name="Arial",size=10,color="333333")
    bfb = Font(name="Arial",size=10,bold=True,color="333333")
    tf = Font(name="Arial",size=14,bold=True,color=PG)
    sf = Font(name="Arial",size=10,color="888888",italic=True)
    rp = "#,##0"
    bdr = Border(left=Side("thin",color="DDDDDD"),right=Side("thin",color="DDDDDD"),
                 top=Side("thin",color="DDDDDD"),bottom=Side("thin",color="DDDDDD"))
    wc = Alignment(wrap_text=True,vertical="center")
    cx = Alignment(horizontal="center",vertical="center")
    gf = PatternFill("solid",fgColor="F5F5F5")
    ch_colors = ["6A4C93","D4A843","457B9D","E76F51","2D6A4F","E63946"]

    def add_hdr(ws,title,sub,cols):
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=cols)
        ws.cell(1,1,title).font=tf; ws.row_dimensions[1].height=28
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=cols)
        ws.cell(2,1,sub).font=sf

    def wr_tbl(ws,sr,hdrs,rows,nc=None):
        r=sr
        for c,h in enumerate(hdrs,1):
            cl=ws.cell(r,c,h);cl.font=hf;cl.fill=hfill;cl.alignment=cx;cl.border=bdr
        for i,rd in enumerate(rows):
            r=sr+1+i
            for c,v in enumerate(rd,1):
                cl=ws.cell(r,c,v);cl.font=bf;cl.border=bdr;cl.alignment=wc
                if nc and c in nc: cl.number_format=rp
                if i%2==0: cl.fill=gf
        return sr+1+len(rows)

    def mk_bar(ws,t,dc,cc,mr1,mr2,anc,colors=None,yt="Nilai",xt=""):
        c=BarChart();c.type="col";c.grouping="clustered";c.style=10
        c.title=t;c.width=22;c.height=12
        if yt: c.y_axis.title=yt
        if xt: c.x_axis.title=xt
        d=Reference(ws,min_col=dc,min_row=mr1,max_row=mr2)
        cats=Reference(ws,min_col=cc,min_row=mr1+1,max_row=mr2)
        c.add_data(d,titles_from_data=True);c.set_categories(cats);c.legend=None
        if colors:
            for i2,cl in enumerate(colors):
                if i2<len(c.series): c.series[i2].graphicalProperties.solidFill=cl
        elif c.series: c.series[0].graphicalProperties.solidFill=PG
        c.y_axis.numFmt=rp
        for s in c.series:
            s.dLbls=DataLabelList();s.dLbls.showVal=True;s.dLbls.numFmt=rp
        ws.add_chart(c,anc)

    def mk_line(ws,t,dcs,cc,mr1,mr2,anc,colors=None,yt="Nilai",xt="Tanggal"):
        c=LineChart();c.style=10;c.title=t;c.width=24;c.height=13
        if yt: c.y_axis.title=yt
        if xt: c.x_axis.title=xt
        cats=Reference(ws,min_col=cc,min_row=mr1+1,max_row=mr2)
        for dc in dcs:
            d=Reference(ws,min_col=dc,min_row=mr1,max_row=mr2)
            c.add_data(d,titles_from_data=True)
        c.set_categories(cats)
        if colors:
            for i2,cl in enumerate(colors):
                if i2<len(c.series):
                    c.series[i2].graphicalProperties.line.solidFill=cl
                    c.series[i2].graphicalProperties.line.width=25000
        c.y_axis.numFmt=rp
        ws.add_chart(c,anc)

    pl = f"Playground TnT  |  Periode: {date_start} s.d. {date_end}"
    df["date_only"]=df["sales_date"].dt.date
    df["weekday"]=df["sales_date"].dt.day_name()
    df["day_type"]=df["sales_date"].apply(lambda d: "Weekend" if d.weekday()>=5 else "Weekday")
    df["week_num"]=df["sales_date"].apply(lambda d:(d.day-1)//7+1)
    df["week_label"]="Week "+df["week_num"].astype(str)
    df["month"]=df["sales_date"].dt.to_period("M").astype(str)

    t_amt=df["amount"].sum();t_nett=df["nett_sales"].sum();t_tax=df["tax_amount"].sum()
    t_trx=len(df);t_child=df["child_total"].sum();t_comp=df["companion_total"].sum()
    n_days=df["date_only"].nunique()

    # Sheet 1: Summary
    ws1=wb.active;ws1.title="Executive Summary";ws1.sheet_properties.tabColor=PG
    for c,w in [(1,28),(2,22)]: ws1.column_dimensions[chr(64+c)].width=w
    add_hdr(ws1,"Playground TnT — Executive Summary",pl,2)
    kpis=[("Total Revenue",t_amt,rp),("Total Nett Sales",t_nett,rp),("Total Tax",t_tax,rp),
          ("Total Transaksi",t_trx,"#,##0"),("Total Anak",t_child,"#,##0"),("Total Pendamping",t_comp,"#,##0"),
          ("Avg/Transaksi",t_amt/t_trx if t_trx else 0,rp),("Avg/Anak",t_amt/t_child if t_child else 0,rp),
          ("Hari Aktif",n_days,"#,##0"),("Avg Daily Revenue",t_amt/n_days if n_days else 0,rp)]
    r=4
    ws1.cell(r,1,"Metrik").font=hf;ws1.cell(r,1).fill=hfill;ws1.cell(r,1).border=bdr
    ws1.cell(r,2,"Nilai").font=hf;ws1.cell(r,2).fill=hfill;ws1.cell(r,2).border=bdr
    for i,(lb,v,fm) in enumerate(kpis):
        r=5+i;ws1.cell(r,1,lb).font=bfb;ws1.cell(r,1).border=bdr
        ws1.cell(r,2,v).font=Font(name="Arial",size=12,bold=True,color=PG)
        ws1.cell(r,2).number_format=fm;ws1.cell(r,2).border=bdr
        if i%2==0: ws1.cell(r,1).fill=gf;ws1.cell(r,2).fill=gf

    # Sheet 2: Trend Harian
    ws2=wb.create_sheet("Trend Harian");ws2.sheet_properties.tabColor=PG
    for c,w in [(1,14),(2,8),(3,15),(4,10),(5,10)]: ws2.column_dimensions[chr(64+c)].width=w
    add_hdr(ws2,"Trend Harian",pl,5)
    daily=df.groupby("date_only").agg(amt=("amount","sum"),trx=("amount","count"),
        child=("child_total","sum"),comp=("companion_total","sum")).reset_index()
    daily["wd"]=pd.to_datetime(daily["date_only"]).dt.strftime("%a")
    dr=[[str(r2["date_only"]),r2["wd"],r2["amt"],r2["child"],r2["trx"]] for _,r2 in daily.iterrows()]
    er=wr_tbl(ws2,4,["Tanggal","Hari","Revenue (Rp)","Anak","Transaksi"],dr,nc={3,4,5})
    mk_bar(ws2,"Daily Revenue",3,1,4,er-1,f"A{er+1}",yt="Revenue (Rp)",xt="Tanggal")
    mk_line(ws2,"Daily Anak",[4],1,4,er-1,f"A{er+16}",colors=[GLD],yt="Jumlah Anak")

    # Sheet 3: Child & Companion
    ws3=wb.create_sheet("Child & Companion");ws3.sheet_properties.tabColor=GLD
    for c,w in [(1,14),(2,12),(3,12)]: ws3.column_dimensions[chr(64+c)].width=w
    add_hdr(ws3,"Child & Companion Analysis",pl,3)
    ws3.cell(4,1,"Tipe").font=hf;ws3.cell(4,1).fill=hfill;ws3.cell(4,1).border=bdr
    ws3.cell(4,2,"Jumlah").font=hf;ws3.cell(4,2).fill=hfill;ws3.cell(4,2).border=bdr
    ws3.cell(5,1,"Anak").font=bf;ws3.cell(5,2,t_child).font=bf;ws3.cell(5,2).number_format=rp
    ws3.cell(6,1,"Pendamping").font=bf;ws3.cell(6,2,t_comp).font=bf;ws3.cell(6,2).number_format=rp
    for r in [4,5,6]:
        for c in [1,2]: ws3.cell(r,c).border=bdr
    pie=PieChart();pie.title="Komposisi Pengunjung";pie.style=10;pie.width=16;pie.height=12
    d=Reference(ws3,min_col=2,min_row=4,max_row=6);cats=Reference(ws3,min_col=1,min_row=5,max_row=6)
    pie.add_data(d,titles_from_data=True);pie.set_categories(cats)
    pie.dataLabels=DataLabelList();pie.dataLabels.showPercent=True;pie.dataLabels.showVal=True;pie.dataLabels.numFmt=rp
    pt0=DataPoint(idx=0);pt0.graphicalProperties.solidFill=PG;pie.series[0].data_points.append(pt0)
    pt1=DataPoint(idx=1);pt1.graphicalProperties.solidFill=GLD;pie.series[0].data_points.append(pt1)
    ws3.add_chart(pie,"A8")
    # Top repeat customers
    freq=df.groupby("customer_name").agg(visits=("order_id","count"),spend=("amount","sum"),ch=("child_total","sum")).reset_index()
    freq=freq.sort_values("visits",ascending=False).head(20)
    r_f=24;ws3.cell(r_f,1,"Top Repeat Customers").font=Font(name="Arial",size=12,bold=True,color=PG)
    r_f+=1
    fr=[[r2["customer_name"],r2["visits"],r2["spend"],r2["ch"]] for _,r2 in freq.iterrows()]
    wr_tbl(ws3,r_f,["Nama","Kunjungan","Total Spend (Rp)","Total Anak"],fr,nc={3,4})

    # Sheet 4: Weekly
    ws4=wb.create_sheet("Weekly Report");ws4.sheet_properties.tabColor="457B9D"
    for c,w in [(1,10),(2,15),(3,10),(4,10)]: ws4.column_dimensions[chr(64+c)].width=w
    add_hdr(ws4,"Weekly Report",pl,4)
    wk=df.groupby("week_label").agg(rev=("amount","sum"),trx=("amount","count"),ch=("child_total","sum")).reset_index()
    wr2=[[r2["week_label"],r2["rev"],r2["trx"],r2["ch"]] for _,r2 in wk.iterrows()]
    er=wr_tbl(ws4,4,["Week","Revenue (Rp)","Transaksi","Anak"],wr2,nc={2,3,4})
    mk_bar(ws4,"Revenue per Week",2,1,4,er-1,f"A{er+1}",yt="Revenue (Rp)",xt="Minggu")

    # Sheet 5: Monthly
    ws5=wb.create_sheet("Monthly Overview");ws5.sheet_properties.tabColor="E76F51"
    for c,w in [(1,12),(2,15),(3,10),(4,10)]: ws5.column_dimensions[chr(64+c)].width=w
    add_hdr(ws5,"Monthly Overview",pl,4)
    mt=df.groupby("month").agg(rev=("amount","sum"),trx=("amount","count"),ch=("child_total","sum")).reset_index().sort_values("month")
    mr2=[[r2["month"],r2["rev"],r2["trx"],r2["ch"]] for _,r2 in mt.iterrows()]
    er=wr_tbl(ws5,4,["Bulan","Revenue (Rp)","Transaksi","Anak"],mr2,nc={2,3,4})
    mk_bar(ws5,"Revenue per Bulan",2,1,4,er-1,f"A{er+1}",yt="Revenue (Rp)",xt="Bulan")
    mk_bar(ws5,"Anak per Bulan",4,1,4,er-1,f"A{er+16}",colors=[GLD],yt="Jumlah Anak",xt="Bulan")

    # Sheet 6: Deep Dive
    ws6=wb.create_sheet("Deep Dive");ws6.sheet_properties.tabColor=PG
    for c,w in [(1,14),(2,15),(3,15),(4,15)]: ws6.column_dimensions[chr(64+c)].width=w
    add_hdr(ws6,"Deep Dive — Moving Average",pl,4)
    dma=df.groupby("date_only")["amount"].sum().reset_index();dma.columns=["Tanggal","Revenue"]
    dma=dma.sort_values("Tanggal")
    dma["MA_7"]=dma["Revenue"].rolling(7,min_periods=1).mean().round(0)
    dma["MA_14"]=dma["Revenue"].rolling(14,min_periods=1).mean().round(0)
    mr2=[[str(r2["Tanggal"]),r2["Revenue"],r2["MA_7"],r2["MA_14"]] for _,r2 in dma.iterrows()]
    er=wr_tbl(ws6,4,["Tanggal","Daily Revenue","MA 7-Day","MA 14-Day"],mr2,nc={2,3,4})
    mk_line(ws6,"Moving Average Analysis",[2,3,4],1,4,er-1,f"A{er+1}",colors=[PGL,PG.replace("4C","72"),"6A4C93"],yt="Revenue (Rp)")

    out=BytesIO();wb.save(out);out.seek(0)
    return out.getvalue()


def generate_playground_html(df, date_start, date_end):
    """Generate Playground HTML report for print/PDF."""
    t_amt=df["amount"].sum();t_trx=len(df);t_child=df["child_total"].sum()
    t_comp=df["companion_total"].sum();n_days=df["sales_date"].dt.date.nunique()
    df["date_only"]=df["sales_date"].dt.date
    df["weekday"]=df["sales_date"].dt.day_name()
    df["day_type"]=df["sales_date"].apply(lambda d:"Weekend" if d.weekday()>=5 else "Weekday")
    df["week_label"]="Week "+df["sales_date"].apply(lambda d:str((d.day-1)//7+1))
    df["month"]=df["sales_date"].dt.to_period("M").astype(str)

    cl = dict(font=dict(family="ui-rounded, Segoe UI, system-ui, sans-serif",size=13),title_font_size=16,
              margin=dict(t=50,b=50,l=60,r=30),height=500)
    parts = []

    def ac(fig, title=None):
        if title: parts.append(f'<div class="section-title">{title}</div>')
        parts.append(fig.to_html(full_html=False,include_plotlyjs=False,config={"displayModeBar":False}))
        parts.append('<div class="page-break"></div>')

    parts.append(f"""
    <div class="kpi-row">
        <div class="kpi" style="border-left-color:#1A6B3F"><div class="kpi-label">Total Revenue</div><div class="kpi-value">Rp {t_amt:,.0f}</div></div>
        <div class="kpi gold"><div class="kpi-label">Transaksi</div><div class="kpi-value">{t_trx:,}</div></div>
        <div class="kpi blue"><div class="kpi-label">Total Anak</div><div class="kpi-value">{t_child:,}</div></div>
        <div class="kpi orange"><div class="kpi-label">Total Pendamping</div><div class="kpi-value">{t_comp:,}</div></div>
        <div class="kpi red"><div class="kpi-label">Avg/Transaksi</div><div class="kpi-value">Rp {t_amt/t_trx:,.0f}</div></div>
    </div>""")

    # Trend
    daily=df.groupby("date_only").agg(amt=("amount","sum"),trx=("amount","count"),ch=("child_total","sum")).reset_index()
    f1=make_subplots(specs=[[{"secondary_y":True}]])
    f1.add_trace(go.Bar(x=daily["date_only"].astype(str),y=daily["amt"],name="Revenue",marker_color="#1A6B3F",opacity=0.85),secondary_y=False)
    f1.add_trace(go.Scatter(x=daily["date_only"].astype(str),y=daily["trx"],name="Transaksi",mode="lines+markers",line=dict(color="#C08A2C",width=2.5)),secondary_y=True)
    f1.update_layout(title="Daily Revenue & Transaction",**cl);ac(f1,"1. Trend Harian")

    # Child vs Companion
    vc=pd.DataFrame({"Tipe":["Anak","Pendamping"],"Jumlah":[t_child,t_comp]})
    f2=px.pie(vc,values="Jumlah",names="Tipe",title="Komposisi Pengunjung",color_discrete_sequence=["#1A6B3F","#C08A2C"],hole=0.4)
    f2.update_layout(**cl);ac(f2,"2. Child & Companion")

    dv=df.groupby("date_only").agg(ch=("child_total","sum"),co=("companion_total","sum")).reset_index()
    f3=go.Figure()
    f3.add_trace(go.Bar(x=dv["date_only"].astype(str),y=dv["ch"],name="Anak",marker_color="#1A6B3F"))
    f3.add_trace(go.Bar(x=dv["date_only"].astype(str),y=dv["co"],name="Pendamping",marker_color="#C08A2C"))
    f3.update_layout(title="Daily Child vs Companion",barmode="stack",**cl);ac(f3)

    # WD/WE
    trf=df.groupby("day_type").agg(rev=("amount","sum"),ch=("child_total","sum")).reset_index()
    f4=px.bar(trf,x="day_type",y="rev",title="Revenue — Weekday vs Weekend",color="day_type",
              color_discrete_map={"Weekday":"#3B4C7A","Weekend":"#C0483C"},text="rev")
    f4.update_traces(textposition="outside",texttemplate="Rp%{text:,.0f}",textfont_size=14)
    f4.update_layout(showlegend=False,**cl);ac(f4,"3. Weekday vs Weekend")

    # Weekly
    wk=df.groupby("week_label").agg(rev=("amount","sum")).reset_index()
    f5=px.bar(wk,x="week_label",y="rev",title="Revenue per Week",text="rev",color_discrete_sequence=["#1A6B3F"])
    f5.update_traces(textposition="outside",texttemplate="Rp%{text:,.0f}",textfont_size=13)
    f5.update_layout(**cl);ac(f5,"4. Weekly Report")

    # Monthly
    mt=df.groupby("month").agg(rev=("amount","sum"),ch=("child_total","sum")).reset_index().sort_values("month")
    f6=px.bar(mt,x="month",y="rev",title="Revenue per Bulan",text="rev",color_discrete_sequence=["#1A6B3F"])
    f6.update_traces(textposition="outside",texttemplate="Rp%{text:,.0f}",textfont_size=13)
    f6.update_layout(**cl);ac(f6,"5. Monthly Overview")

    # MA
    dma=df.groupby("date_only")["amount"].sum().reset_index();dma.columns=["T","R"];dma=dma.sort_values("T")
    dma["MA7"]=dma["R"].rolling(7,min_periods=1).mean();dma["MA14"]=dma["R"].rolling(14,min_periods=1).mean()
    f7=go.Figure()
    f7.add_trace(go.Scatter(x=dma["T"].astype(str),y=dma["R"],name="Daily",mode="lines",line=dict(color="#DCEEE4",width=1)))
    f7.add_trace(go.Scatter(x=dma["T"].astype(str),y=dma["MA7"],name="MA 7",mode="lines",line=dict(color="#2E9160",width=2.5)))
    f7.add_trace(go.Scatter(x=dma["T"].astype(str),y=dma["MA14"],name="MA 14",mode="lines",line=dict(color="#1A6B3F",width=2.5)))
    f7.update_layout(title="Moving Average Analysis",**cl);ac(f7,"6. Deep Dive")

    # Summary table
    freq=df.groupby("customer_name").agg(v=("order_id","count"),s=("amount","sum")).reset_index().sort_values("v",ascending=False).head(15)
    tbl='<div class="section-title">Top 15 Repeat Customers</div><table><tr><th>Nama</th><th>Kunjungan</th><th>Total Spend</th></tr>'
    for _,r2 in freq.iterrows():
        tbl+=f'<tr><td>{r2["customer_name"]}</td><td>{r2["v"]}</td><td>Rp {r2["s"]:,.0f}</td></tr>'
    tbl+='</table>';parts.append(tbl)

    body="\\n".join(parts)
    html=f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>GROVE — Playground TnT Report</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');
*{{font-family:ui-rounded, 'Segoe UI', system-ui, -apple-system, sans-serif;margin:0;padding:0;box-sizing:border-box;}}
body{{background:#fff;padding:30px 40px;color:#333;max-width:1100px;margin:0 auto;}}
.report-header{{background:linear-gradient(135deg,#1A6B3F,#2E9160,#C9A9E9);color:white;padding:24px 32px;border-radius:12px;margin-bottom:24px;}}
.report-header h1{{font-size:28px;font-weight:800;}}.report-header p{{font-size:14px;opacity:0.85;}}
.kpi-row{{display:flex;gap:12px;margin-bottom:24px;flex-wrap:wrap;}}
.kpi{{flex:1;min-width:160px;background:#fff;border-radius:10px;padding:16px;border-left:5px solid #1A6B3F;box-shadow:0 1px 4px rgba(0,0,0,0.06);}}
.kpi.gold{{border-left-color:#C08A2C;}}.kpi.blue{{border-left-color:#3B4C7A;}}.kpi.red{{border-left-color:#C0483C;}}.kpi.orange{{border-left-color:#9E6B4A;}}
.kpi-label{{font-size:11px;color:#666;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;}}
.kpi-value{{font-size:22px;font-weight:800;color:#1A6B3F;margin-top:4px;}}
.section-title{{font-size:20px;font-weight:700;color:#1A6B3F;border-bottom:3px solid #2E9160;padding-bottom:6px;margin:32px 0 16px;}}
table{{width:100%;border-collapse:collapse;margin:16px 0;font-size:13px;}}
th{{background:#1A6B3F;color:white;padding:10px 12px;text-align:left;}}td{{padding:8px 12px;border-bottom:1px solid #eee;}}
tr:nth-child(even){{background:#f9f9f9;}}.page-break{{page-break-after:auto;}}
.print-btn{{position:fixed;top:16px;right:16px;z-index:999;background:#1A6B3F;color:white;border:none;padding:12px 24px;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,0.2);}}
@media print{{.print-btn{{display:none;}}body{{padding:10px;}}.report-header,.kpi,th{{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}.page-break{{page-break-after:always;}}}}
</style></head><body>
<button class="print-btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
<div class="report-header"><h1>🎪 Playground TnT Report</h1><p>Periode: {date_start} s.d. {date_end} | Generated: {datetime.now().strftime("%d %B %Y, %H:%M")}</p></div>
{body}
<div style="text-align:center;margin-top:40px;padding:16px;color:#888;font-size:12px;border-top:1px solid #eee;">GROVE Sales Analytics v{APP_VERSION}</div>
</body></html>"""
    return html.encode("utf-8")


def generate_master_xlsx(df_fnb, df_pg, date_start, date_end):
    """Generate Master Dashboard XLSX."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint

    wb = Workbook()
    GD="1B4332";PG="6A4C93";GLD="D4A843"
    hf=Font(name="Arial",size=11,bold=True,color="FFFFFF")
    hfill=PatternFill("solid",fgColor=GD)
    bf=Font(name="Arial",size=10,color="333333")
    bfb=Font(name="Arial",size=10,bold=True,color="333333")
    tf=Font(name="Arial",size=14,bold=True,color=GD)
    sf=Font(name="Arial",size=10,color="888888",italic=True)
    rp="#,##0"
    bdr=Border(left=Side("thin",color="DDDDDD"),right=Side("thin",color="DDDDDD"),
               top=Side("thin",color="DDDDDD"),bottom=Side("thin",color="DDDDDD"))
    wc=Alignment(wrap_text=True,vertical="center");cx=Alignment(horizontal="center",vertical="center")
    gf=PatternFill("solid",fgColor="F5F5F5")

    def wr_tbl(ws,sr,hdrs,rows,nc=None):
        r=sr
        for c,h in enumerate(hdrs,1):
            cl=ws.cell(r,c,h);cl.font=hf;cl.fill=hfill;cl.alignment=cx;cl.border=bdr
        for i,rd in enumerate(rows):
            r=sr+1+i
            for c,v in enumerate(rd,1):
                cl=ws.cell(r,c,v);cl.font=bf;cl.border=bdr;cl.alignment=wc
                if nc and c in nc: cl.number_format=rp
                if i%2==0: cl.fill=gf
        return sr+1+len(rows)

    fnb_nett=df_fnb["nett_sales"].sum() if not df_fnb.empty else 0
    pg_amt=df_pg["amount"].sum() if not df_pg.empty else 0
    grand=fnb_nett+pg_amt
    pl=f"Consolidated  |  Periode: {date_start} s.d. {date_end}"

    # Sheet 1: Summary
    ws1=wb.active;ws1.title="Master Summary";ws1.sheet_properties.tabColor=GD
    for c,w in [(1,28),(2,22)]: ws1.column_dimensions[chr(64+c)].width=w
    ws1.merge_cells("A1:B1");ws1.cell(1,1,"GROVE Sales Analytics — Master Summary").font=tf
    ws1.merge_cells("A2:B2");ws1.cell(2,1,pl).font=sf

    kpis=[("Grand Total Revenue",grand,rp),("F&B Nett Sales",fnb_nett,rp),("Playground Revenue",pg_amt,rp),
          ("F&B Pax",df_fnb["pax_total"].sum() if not df_fnb.empty else 0,"#,##0"),
          ("Playground Anak",df_pg["child_total"].sum() if not df_pg.empty else 0,"#,##0"),
          ("F&B Kontribusi",fnb_nett/grand if grand else 0,"0.0%"),
          ("Playground Kontribusi",pg_amt/grand if grand else 0,"0.0%")]
    r=4
    ws1.cell(r,1,"Metrik").font=hf;ws1.cell(r,1).fill=hfill;ws1.cell(r,1).border=bdr
    ws1.cell(r,2,"Nilai").font=hf;ws1.cell(r,2).fill=hfill;ws1.cell(r,2).border=bdr
    for i,(lb,v,fm) in enumerate(kpis):
        r=5+i;ws1.cell(r,1,lb).font=bfb;ws1.cell(r,1).border=bdr
        ws1.cell(r,2,v).font=Font(name="Arial",size=12,bold=True,color=GD)
        ws1.cell(r,2).number_format=fm;ws1.cell(r,2).border=bdr
        if i%2==0: ws1.cell(r,1).fill=gf;ws1.cell(r,2).fill=gf

    # Contribution pie
    r_p=r+2;ws1.cell(r_p,1,"Segment").font=hf;ws1.cell(r_p,1).fill=hfill;ws1.cell(r_p,1).border=bdr
    ws1.cell(r_p,2,"Revenue").font=hf;ws1.cell(r_p,2).fill=hfill;ws1.cell(r_p,2).border=bdr
    ws1.cell(r_p+1,1,"F&B Tenants").font=bf;ws1.cell(r_p+1,2,fnb_nett).font=bf;ws1.cell(r_p+1,2).number_format=rp
    ws1.cell(r_p+2,1,"Playground TnT").font=bf;ws1.cell(r_p+2,2,pg_amt).font=bf;ws1.cell(r_p+2,2).number_format=rp
    for rx in [r_p,r_p+1,r_p+2]:
        for c in [1,2]: ws1.cell(rx,c).border=bdr
    pie=PieChart();pie.title="Kontribusi Revenue";pie.style=10;pie.width=16;pie.height=12
    d=Reference(ws1,min_col=2,min_row=r_p,max_row=r_p+2);cats=Reference(ws1,min_col=1,min_row=r_p+1,max_row=r_p+2)
    pie.add_data(d,titles_from_data=True);pie.set_categories(cats)
    pie.dataLabels=DataLabelList();pie.dataLabels.showPercent=True;pie.dataLabels.showVal=True;pie.dataLabels.numFmt=rp
    pt0=DataPoint(idx=0);pt0.graphicalProperties.solidFill=GD;pie.series[0].data_points.append(pt0)
    pt1=DataPoint(idx=1);pt1.graphicalProperties.solidFill=PG;pie.series[0].data_points.append(pt1)
    ws1.add_chart(pie,f"A{r_p+4}")

    # Sheet 2: Monthly comparison
    ws2=wb.create_sheet("Monthly Comparison");ws2.sheet_properties.tabColor=GLD
    for c,w in [(1,12),(2,18),(3,18),(4,18)]: ws2.column_dimensions[chr(64+c)].width=w
    ws2.merge_cells("A1:D1");ws2.cell(1,1,"Monthly Revenue Comparison — F&B vs Playground").font=tf
    ws2.merge_cells("A2:D2");ws2.cell(2,1,pl).font=sf

    mt_fnb=df_fnb.groupby(df_fnb["sales_date"].dt.to_period("M").astype(str))["nett_sales"].sum().reset_index() if not df_fnb.empty else pd.DataFrame(columns=["sales_date","nett_sales"])
    mt_fnb.columns=["month","F&B"]
    mt_pg=df_pg.groupby(df_pg["sales_date"].dt.to_period("M").astype(str))["amount"].sum().reset_index() if not df_pg.empty else pd.DataFrame(columns=["sales_date","amount"])
    mt_pg.columns=["month","Playground"]
    mt=pd.merge(mt_fnb,mt_pg,on="month",how="outer").fillna(0).sort_values("month")
    mt["Grand Total"]=mt["F&B"]+mt["Playground"]
    mr=[[r2["month"],r2["F&B"],r2["Playground"],r2["Grand Total"]] for _,r2 in mt.iterrows()]
    er=wr_tbl(ws2,4,["Bulan","F&B (Rp)","Playground (Rp)","Grand Total (Rp)"],mr,nc={2,3,4})

    # Grouped bar
    ch=BarChart();ch.type="col";ch.grouping="clustered";ch.style=10;ch.title="Monthly Revenue — F&B vs Playground"
    ch.width=24;ch.height=13;ch.y_axis.title="Revenue (Rp)";ch.x_axis.title="Bulan"
    for dc,clr in [(2,GD),(3,PG)]:
        d=Reference(ws2,min_col=dc,min_row=4,max_row=er-1)
        ch.add_data(d,titles_from_data=True)
    cats=Reference(ws2,min_col=1,min_row=5,max_row=er-1);ch.set_categories(cats)
    if len(ch.series)>0: ch.series[0].graphicalProperties.solidFill=GD
    if len(ch.series)>1: ch.series[1].graphicalProperties.solidFill=PG
    ch.y_axis.numFmt=rp
    for s in ch.series:
        s.dLbls=DataLabelList();s.dLbls.showVal=True;s.dLbls.numFmt=rp
    ws2.add_chart(ch,f"A{er+1}")

    out=BytesIO();wb.save(out);out.seek(0)
    return out.getvalue()


def generate_master_html(df_fnb, df_pg, date_start, date_end):
    """Generate Master Dashboard HTML for print/PDF."""
    fnb_nett=df_fnb["nett_sales"].sum() if not df_fnb.empty else 0
    pg_amt=df_pg["amount"].sum() if not df_pg.empty else 0
    fnb_pax=df_fnb["pax_total"].sum() if not df_fnb.empty else 0
    pg_child=df_pg["child_total"].sum() if not df_pg.empty else 0
    grand=fnb_nett+pg_amt
    cl=dict(font=dict(family="ui-rounded, Segoe UI, system-ui, sans-serif",size=13),title_font_size=16,margin=dict(t=50,b=50,l=60,r=30),height=500)
    parts=[]
    def ac(fig,t=None):
        if t: parts.append(f'<div class="section-title">{t}</div>')
        parts.append(fig.to_html(full_html=False,include_plotlyjs=False,config={"displayModeBar":False}))
        parts.append('<div class="page-break"></div>')

    parts.append(f"""
    <div class="kpi-row">
        <div class="kpi"><div class="kpi-label">Grand Total Revenue</div><div class="kpi-value">Rp {grand:,.0f}</div></div>
        <div class="kpi gold"><div class="kpi-label">F&B Nett Sales</div><div class="kpi-value">Rp {fnb_nett:,.0f}</div></div>
        <div class="kpi" style="border-left-color:#1A6B3F"><div class="kpi-label">Playground Revenue</div><div class="kpi-value">Rp {pg_amt:,.0f}</div></div>
        <div class="kpi orange"><div class="kpi-label">F&B Pax</div><div class="kpi-value">{fnb_pax:,}</div></div>
        <div class="kpi red"><div class="kpi-label">Playground Anak</div><div class="kpi-value">{pg_child:,}</div></div>
    </div>""")

    # Pie
    ct=pd.DataFrame({"Segment":["F&B Tenants","Playground TnT"],"Revenue":[fnb_nett,pg_amt]})
    fp=px.pie(ct,values="Revenue",names="Segment",title="Kontribusi Revenue",color_discrete_sequence=["#1A6B3F","#1A6B3F"],hole=0.4)
    fp.update_layout(**cl);ac(fp,"1. Revenue Overview")

    # Daily stacked
    d_fnb=df_fnb.groupby(df_fnb["sales_date"].dt.date)["nett_sales"].sum().reset_index() if not df_fnb.empty else pd.DataFrame(columns=["sales_date","nett_sales"])
    d_fnb.columns=["date","F&B"]
    d_pg=df_pg.groupby(df_pg["sales_date"].dt.date)["amount"].sum().reset_index() if not df_pg.empty else pd.DataFrame(columns=["sales_date","amount"])
    d_pg.columns=["date","Playground"]
    mg=pd.merge(d_fnb,d_pg,on="date",how="outer").fillna(0).sort_values("date")
    fd=go.Figure()
    fd.add_trace(go.Bar(x=mg["date"].astype(str),y=mg["F&B"],name="F&B",marker_color="#1A6B3F"))
    fd.add_trace(go.Bar(x=mg["date"].astype(str),y=mg["Playground"],name="Playground",marker_color="#1A6B3F"))
    fd.update_layout(title="Daily Revenue — F&B vs Playground",barmode="stack",**cl);ac(fd,"2. Daily Comparison")

    # Monthly
    mt_fnb=df_fnb.groupby(df_fnb["sales_date"].dt.to_period("M").astype(str))["nett_sales"].sum().reset_index() if not df_fnb.empty else pd.DataFrame(columns=["sales_date","nett_sales"])
    mt_fnb.columns=["month","F&B"]
    mt_pg=df_pg.groupby(df_pg["sales_date"].dt.to_period("M").astype(str))["amount"].sum().reset_index() if not df_pg.empty else pd.DataFrame(columns=["sales_date","amount"])
    mt_pg.columns=["month","Playground"]
    mt=pd.merge(mt_fnb,mt_pg,on="month",how="outer").fillna(0).sort_values("month")
    fm=go.Figure()
    fm.add_trace(go.Bar(x=mt["month"],y=mt["F&B"],name="F&B",marker_color="#1A6B3F"))
    fm.add_trace(go.Bar(x=mt["month"],y=mt["Playground"],name="Playground",marker_color="#1A6B3F"))
    fm.update_layout(title="Monthly Revenue — F&B vs Playground",barmode="group",**cl);ac(fm,"3. Monthly Comparison")

    # Table
    mt["Grand Total"]=mt["F&B"]+mt["Playground"]
    tbl='<div class="section-title">Ringkasan Bulanan</div><table><tr><th>Bulan</th><th>F&B</th><th>Playground</th><th>Grand Total</th></tr>'
    for _,r2 in mt.iterrows():
        tbl+=f'<tr><td>{r2["month"]}</td><td>Rp {r2["F&B"]:,.0f}</td><td>Rp {r2["Playground"]:,.0f}</td><td>Rp {r2["Grand Total"]:,.0f}</td></tr>'
    tbl+='</table>';parts.append(tbl)

    body="\\n".join(parts)
    html=f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>GROVE — Master Dashboard Report</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');
*{{font-family:ui-rounded, 'Segoe UI', system-ui, -apple-system, sans-serif;margin:0;padding:0;box-sizing:border-box;}}
body{{background:#fff;padding:30px 40px;color:#333;max-width:1100px;margin:0 auto;}}
.report-header{{background:linear-gradient(135deg,#0B140F,#103D28,#C08A2C);color:white;padding:24px 32px;border-radius:12px;margin-bottom:24px;}}
.report-header h1{{font-size:28px;font-weight:800;}}.report-header p{{font-size:14px;opacity:0.85;}}
.kpi-row{{display:flex;gap:12px;margin-bottom:24px;flex-wrap:wrap;}}
.kpi{{flex:1;min-width:160px;background:#fff;border-radius:10px;padding:16px;border-left:5px solid #103D28;box-shadow:0 1px 4px rgba(0,0,0,0.06);}}
.kpi.gold{{border-left-color:#C08A2C;}}.kpi.orange{{border-left-color:#9E6B4A;}}.kpi.red{{border-left-color:#C0483C;}}
.kpi-label{{font-size:11px;color:#666;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;}}
.kpi-value{{font-size:22px;font-weight:800;color:#103D28;margin-top:4px;}}
.section-title{{font-size:20px;font-weight:700;color:#103D28;border-bottom:3px solid #C08A2C;padding-bottom:6px;margin:32px 0 16px;}}
table{{width:100%;border-collapse:collapse;margin:16px 0;font-size:13px;}}
th{{background:#103D28;color:white;padding:10px 12px;text-align:left;}}td{{padding:8px 12px;border-bottom:1px solid #eee;}}
tr:nth-child(even){{background:#f9f9f9;}}.page-break{{page-break-after:auto;}}
.print-btn{{position:fixed;top:16px;right:16px;z-index:999;background:#103D28;color:white;border:none;padding:12px 24px;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;}}
@media print{{.print-btn{{display:none;}}body{{padding:10px;}}.report-header,.kpi,th{{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}.page-break{{page-break-after:always;}}}}
</style></head><body>
<button class="print-btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
<div class="report-header"><h1>🏠 GROVE Master Dashboard Report</h1><p>Consolidated F&B + Playground | Periode: {date_start} s.d. {date_end} | Generated: {datetime.now().strftime("%d %B %Y, %H:%M")}</p></div>
{body}
<div style="text-align:center;margin-top:40px;padding:16px;color:#888;font-size:12px;border-top:1px solid #eee;">GROVE Sales Analytics v{APP_VERSION}</div>
</body></html>"""
    return html.encode("utf-8")


# ============================================================================
# DASHBOARD
# ============================================================================
# ============================================================================
# SHARED DASHBOARD SECTIONS
# ----------------------------------------------------------------------------
# Four of the six sections on each dashboard are the same analysis over
# different measures -- F&B counts pax, Playground counts children plus
# companions. Building them once, parameterised by column name, keeps the two
# pages from drifting apart visually and halves the surface where a bug can
# hide.
# ============================================================================
GREEN, GREEN_D, OCHRE, LINE_C = "#1A6B3F", "#103D28", "#C08A2C", "#EBECE8"
FLAT = "#F3F4F1"


def _daily(d, val, cnt):
    g = d.groupby(d["sales_date"].dt.date).agg(v=(val, "sum"), c=(cnt, "sum")).reset_index()
    g.columns = ["day", "v", "c"]
    # Back to datetime64: a column of datetime.date objects can reach Plotly
    # unparsed, and an axis with nothing it recognises falls back to the
    # current clock -- which is why the trend chart showed seconds instead of
    # a month of dates.
    g["day"] = pd.to_datetime(g["day"], errors="coerce")
    g = g.dropna(subset=["day"]).sort_values("day")
    return g


def _bars_vs_avg(x, y, height=210, hatch_below=True):
    """Bars that carry their own threshold: solid at or above the mean of the
    trading periods, hatched below it. Empty periods stay flat rather than
    drawn as stubs, which would read as a real but tiny figure."""
    act = [v for v in y if v > 0]
    avg = sum(act) / len(act) if act else 0
    fig = go.Figure(go.Bar(
        x=x, y=y,
        marker=dict(color=[GREEN if v >= avg and v > 0 else FLAT for v in y],
                    pattern=dict(shape=["" if (v >= avg and v > 0) else ("/" if hatch_below else "")
                                        for v in y],
                                 fgcolor="#D9DDD7", size=5, solidity=.3)),
        hovertemplate="%{x}<br>Rp %{y:,.0f}<extra></extra>"))
    fig.update_layout(height=height, showlegend=False, bargap=.34,
                      yaxis=dict(visible=False), margin=dict(t=10, b=34, l=10, r=10))
    return fig


def sec_trend(d, cfg):
    g = _daily(d, cfg["val"], cfg["cnt"])
    if g.empty or not g["v"].notna().any():
        st.info("Tidak ada data pada rentang ini."); return

    c = st.columns([2, 1])
    with c[0]:
        with st.container(border=True):
            st.markdown("###### Tren penjualan harian")
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=g["day"], y=g["v"], mode="lines", line=dict(color=GREEN, width=2),
                fill="tozeroy", fillcolor="rgba(26,107,63,.10)",
                hovertemplate="%{x|%d %b}<br>Rp %{y:,.0f}<extra></extra>"))
            peak = g.loc[g["v"].idxmax()]
            fig.add_trace(go.Scatter(
                x=[peak["day"]], y=[peak["v"]], mode="markers",
                marker=dict(color=GREEN, size=9), hoverinfo="skip", showlegend=False))
            fig.update_layout(height=250, showlegend=False,
                              margin=dict(t=10, b=34, l=54, r=14))
            show_chart(fig)

    with c[1]:
        best, worst = g.loc[g["v"].idxmax()], g.loc[g["v"].idxmin()]
        span = (g["day"].max() - g["day"].min()).days + 1
        render_card("Ringkasan periode", pd.DataFrame([
            {"Ukuran": "Hari tertinggi", "Keterangan": f"{best['day']:%d %b %Y}",
             "Nilai": fmt_rp(best["v"])},
            {"Ukuran": "Hari terendah", "Keterangan": f"{worst['day']:%d %b %Y}",
             "Nilai": fmt_rp(worst["v"])},
            {"Ukuran": "Rata-rata harian",
             "Keterangan": f"{len(g)} hari berdagang dari {span} hari",
             "Nilai": fmt_rp(g["v"].mean())}]),
            foot="Hari tanpa transaksi tidak ikut menurunkan rata-rata.")

    days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    dow = (d.assign(w=d["sales_date"].dt.dayofweek)
             .groupby("w")[cfg["val"]].mean().reindex(range(7), fill_value=0))
    with st.container(border=True):
        st.markdown("###### Rata-rata per hari dalam seminggu")
        show_chart(_bars_vs_avg(days, [float(dow.get(i, 0)) for i in range(7)]))


def sec_weekly(d, cfg):
    w = d.assign(wk=d["sales_date"].dt.to_period("W").apply(lambda p: p.start_time))
    g = w.groupby("wk").agg(v=(cfg["val"], "sum"), c=(cfg["cnt"], "sum")).reset_index()
    if g.empty:
        st.info("Tidak ada data."); return
    g["label"] = g["wk"].dt.strftime("%d %b")
    g["wow"] = g["v"].pct_change() * 100
    g["per"] = (g["v"] / g["c"].replace(0, pd.NA)).fillna(0).round()

    with st.container(border=True):
        st.markdown("###### Penjualan per minggu")
        show_chart(_bars_vs_avg(g["label"].tolist(), g["v"].tolist(), height=230))

    render_card("Perubahan antar minggu",
        rows_html([(f"W{i}", f"Minggu {r['label']}",
                    f"{fmt_rp(r['per'])} {cfg['per_label']} · {r['c']:,.0f} {cfg['cnt_label'].lower()}"
                    .replace(",", "."),
                    move_tag(None if pd.isna(r["wow"]) else r["wow"]))
                   for i, (_, r) in enumerate(g.iterrows(), 1)]),
        foot="Minggu pertama tidak punya pembanding, jadi ditandai sebagai baru.")


def sec_monthly(d, cfg):
    m = d.assign(mo=d["sales_date"].dt.to_period("M").astype(str))
    g = m.groupby("mo").agg(v=(cfg["val"], "sum"), c=(cfg["cnt"], "sum"),
                            days=("sales_date", lambda s: s.dt.date.nunique())).reset_index()
    if g.empty:
        st.info("Tidak ada data."); return
    g["mom"] = g["v"].pct_change() * 100
    g["per_day"] = g["v"] / g["days"].replace(0, pd.NA)

    c = st.columns([2, 1])
    with c[0]:
        with st.container(border=True):
            st.markdown("###### Penjualan per bulan")
            show_chart(_bars_vs_avg(g["mo"].tolist(), g["v"].tolist(), height=240))
    with c[1]:
        render_card("Perbandingan bulan",
            rows_html([(r["mo"], r["mo"],
                        f"{r['days']} hari · {fmt_rp(r['per_day'])} / hari",
                        move_tag(None if pd.isna(r["mom"]) else r["mom"]))
                       for _, r in g.iterrows()]),
            foot="Rata-rata harian membuat bulan pendek tetap sebanding dengan bulan penuh.")


def sec_deep(d, cfg, extra_cols=None):
    g = _daily(d, cfg["val"], cfg["cnt"])
    g["per"] = (g["v"] / g["c"].replace(0, pd.NA)).fillna(0).round()
    g["day"] = pd.to_datetime(g["day"])
    with st.container(border=True):
        st.markdown("###### Rincian harian")
        st.dataframe(
            g.rename(columns={"day": "Tanggal", "v": "Nett Sales (Rp)",
                              "c": cfg["cnt_label"], "per": f"Rata {cfg['per_label']} (Rp)"}),
            use_container_width=True, hide_index=True, height=420,
            column_config={
                "Tanggal": st.column_config.DateColumn("Tanggal", format="ddd, DD MMM YYYY"),
                "Nett Sales (Rp)": st.column_config.NumberColumn(format="localized"),
                cfg["cnt_label"]: st.column_config.NumberColumn(format="localized"),
                f"Rata {cfg['per_label']} (Rp)": st.column_config.NumberColumn(format="localized"),
            })
    if extra_cols:
        render_card("Catatan", extra_cols)


def _filters(df, key):
    dmin, dmax = df["sales_date"].min().date(), df["sales_date"].max().date()
    dr = st.date_input("Rentang tanggal",
                       value=(max(dmin, dmax - timedelta(days=30)), dmax),
                       min_value=dmin, max_value=dmax, key=key,
                       label_visibility="collapsed")
    if isinstance(dr, tuple) and len(dr) == 2:
        s, e = pd.Timestamp(dr[0]), pd.Timestamp(dr[1])
        return df[(df["sales_date"] >= s) & (df["sales_date"] <= e)], dr
    return df, (dmin, dmax)


def page_dashboard_fnb():
    render_header()
    db = get_db()
    user = st.session_state["user"]
    tenant_access = user.get("tenant_access", "ALL")
    tenants_df = db.get_tenants()
    if tenant_access == "ALL":
        tenant_list = ["All"] + (tenants_df["tenant_name"].tolist() if not tenants_df.empty else [])
    else:
        tenant_list = [t.strip() for t in tenant_access.split(",")]

    hc = st.columns([2, 1, 1])
    with hc[1]:
        sel_tenant = st.selectbox("Tenant", tenant_list, label_visibility="collapsed")
    df_all = db.get_sales_data(None if sel_tenant == "All" else sel_tenant)
    if df_all.empty:
        with hc[0]:
            render_page_head("Dashboard F&B", "Belum ada data")
        render_card("Belum ada data",
                    '<p style="font-size:12.5px;color:#8B948D;margin:0">Upload file ESB '
                    'melalui menu Upload F&amp;B.</p>')
        return
    with hc[2]:
        df, dr = _filters(df_all, "fnb_dr")
    with hc[0]:
        render_page_head("Dashboard F&B",
                         f"{sel_tenant} · {dr[0]:%d %b %Y} – {dr[1]:%d %b %Y}")
    if df.empty:
        render_card("Tidak ada data",
                    '<p style="font-size:12.5px;color:#8B948D;margin:0">Rentang tanggal '
                    'yang dipilih tidak memuat transaksi.</p>')
        return

    df = enrich_df(df)
    cfg = dict(val="nett_sales", cnt="pax_total", cnt_label="Pax", per_label="per pax")

    nett, pax = df["nett_sales"].sum(), df["pax_total"].sum()
    disc, sub = df["discount_total"].sum(), df["subtotal"].sum()
    ndays = df["sales_date"].dt.date.nunique()
    k = st.columns(4)
    with k[0]: render_kpi("Nett Sales", fmt_rp(nett), featured=True,
                          caption=f"{ndays} hari berdagang")
    with k[1]: render_kpi("Pax", f"{pax:,.0f}".replace(",", "."),
                          caption=f"{pax/ndays:,.0f} per hari".replace(",", ".") if ndays else None)
    with k[2]: render_kpi("Rata per Pax", fmt_rp(nett / pax if pax else 0),
                          caption="nett dibagi jumlah pax")
    with k[3]: render_kpi("Diskon", fmt_rp(disc),
                          caption=f"{disc/sub*100:.1f}% dari subtotal" if sub else None)

    t = st.tabs(["Trend Harian", "Analisis Per Jam", "Time Segment",
                 "Weekly Report", "Monthly Overview", "Deep Dive"])

    with t[0]:
        sec_trend(df, cfg)

    with t[1]:
        hourly = df.groupby("hr").agg(v=("nett_sales", "sum"), p=("pax_total", "sum")).reset_index()
        c = st.columns([2, 1])
        with c[0]:
            with st.container(border=True):
                st.markdown("###### Penjualan per jam")
                show_chart(_bars_vs_avg([f"{h:02d}" for h in hourly["hr"]],
                                        hourly["v"].tolist(), height=240))
        with c[1]:
            top = hourly.nlargest(5, "v")
            render_card("Jam tersibuk",
                rows_html([(f"H{int(r['hr'])}", f"Pukul {int(r['hr']):02d}:00",
                            f"{r['p']:,.0f} pax".replace(",", "."),
                            f'<span class="amt num">{fmt_rp(r["v"])}</span>')
                           for _, r in top.iterrows()]),
                foot="Lima jam dengan penjualan tertinggi pada rentang ini.")
        heat = (df.assign(w=df["sales_date"].dt.dayofweek)
                  .pivot_table(index="w", columns="hr", values="nett_sales", aggfunc="sum")
                  .reindex(range(7)).fillna(0))
        with st.container(border=True):
            st.markdown("###### Peta jam × hari")
            fig = go.Figure(go.Heatmap(
                z=heat.values, x=[f"{int(c):02d}" for c in heat.columns],
                y=["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"],
                colorscale=[[0, "#F5F7F4"], [.5, "#7FB79A"], [1, GREEN_D]],
                hovertemplate="%{y} %{x}:00<br>Rp %{z:,.0f}<extra></extra>", showscale=False))
            fig.update_layout(height=230, margin=dict(t=10, b=34, l=54, r=14))
            show_chart(fig)

    with t[2]:
        seg = df[df["segment"] != "Other"]
        if seg.empty:
            st.info("Tidak ada transaksi pada jam segmen yang ditentukan.")
        else:
            agg = seg.groupby("segment").agg(v=("nett_sales", "sum"),
                                             p=("pax_total", "sum")).reset_index()
            c = st.columns([1, 2])
            with c[0]:
                with st.container(border=True):
                    st.markdown("###### Kontribusi segmen")
                    fig = go.Figure(go.Pie(labels=agg["segment"], values=agg["v"], hole=.62,
                        marker=dict(colors=[GREEN, OCHRE, "#6B7B3A"], line=dict(color="#fff", width=2)),
                        textinfo="percent", hovertemplate="%{label}<br>Rp %{value:,.0f}<extra></extra>"))
                    fig.update_layout(height=250, margin=dict(t=10, b=10, l=10, r=10))
                    show_chart(fig)
            with c[1]:
                agg["per"] = (agg["v"] / agg["p"].replace(0, pd.NA)).fillna(0).round()
                render_card("Perbandingan segmen",
                    rows_html([(r["segment"], r["segment"],
                                f"{r['p']:,.0f} pax · {fmt_rp(r['per'])} per pax".replace(",", "."),
                                f'<span class="amt num">{fmt_rp(r["v"])}</span>')
                               for _, r in agg.sort_values("v", ascending=False).iterrows()]),
                    foot="Breakfast 07–10, Lunch 12–14, After Office 17–19.")

    with t[3]: sec_weekly(df, cfg)
    with t[4]: sec_monthly(df, cfg)
    with t[5]:
        sec_deep(df, cfg)
        if df["tenant_name"].nunique() > 1:
            per_t = (df.groupby(["tenant_id", "tenant_name"])
                       .agg(v=("nett_sales", "sum"), p=("pax_total", "sum")).reset_index()
                       .sort_values("v", ascending=False))
            render_card("Perbandingan tenant",
                rows_html([(r["tenant_id"], r["tenant_name"],
                            f"{r['p']:,.0f} pax".replace(",", "."),
                            f'<span class="amt num">{fmt_rp(r["v"])}</span>')
                           for _, r in per_t.iterrows()]))

    _export_row(df, sel_tenant, dr, kind="fnb")


def page_upload_esb():
    render_header()
    db = get_db()
    user = st.session_state["user"]

    st.subheader("📤 Upload Data ESB")
    st.markdown("""
    <div class="upload-info">
        <strong>Petunjuk:</strong> Upload file Excel output dari ESB POS
        (<em>Sales Recapitulation by Time Report</em>). Sistem akan otomatis mendeteksi
        nama tenant dari kolom Branch Name dan menambahkan data baru tanpa duplikasi.<br>
        <strong>💡 Bisa upload beberapa file sekaligus!</strong>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader("Pilih file ESB (.xlsx)", type=["xlsx"], accept_multiple_files=True)

    if uploaded:
        for file in uploaded:
            with st.expander(f"📄 {file.name}", expanded=True):
                try:
                    raw = file.read()
                    rows, meta = ESBParser.parse(raw, file.name)
                    if not rows:
                        st.error("File tidak berisi data valid.")
                        continue

                    branches = {}
                    for r in rows:
                        branches.setdefault(r["branch_name"], []).append(r)

                    for branch, branch_rows in branches.items():
                        tenant = db.get_tenant_by_branch(branch)
                        if not tenant:
                            st.warning(f"⚠️ Branch **{branch}** belum terdaftar.")
                            ca, cb = st.columns([3,1])
                            with ca:
                                new_name = st.text_input(f"Nama tenant untuk '{branch}':", value=branch,
                                                         key=f"n_{branch}_{file.name}")
                            with cb:
                                st.markdown("<br>", unsafe_allow_html=True)
                                if st.button("➕ Daftarkan", key=f"r_{branch}_{file.name}"):
                                    db.add_tenant(new_name, branch)
                                    st.success(f"✅ Tenant '{new_name}' terdaftar.")
                                    st.rerun()
                            continue

                        tenant_name = tenant["tenant_name"]
                        tenant_id = tenant["tenant_id"]
                        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        insert_rows = [
                            (tenant_id, r["sales_date"], r["sales_hour"],
                             r["pax_total"], r["subtotal"], r["discount_total"],
                             r["nett_sales"], user["email"], now)
                            for r in branch_rows
                        ]

                        st.markdown(f"**Tenant:** {tenant_name} · **Data:** {len(branch_rows)} baris")

                        # Say what a repeat upload will do before it happens.
                        # The Sheets version appended blindly, so re-uploading a
                        # day doubled it; here the slots are rewritten, but the
                        # user should still see how many are affected.
                        overlap = db.count_existing_slots(tenant_id, branch_rows)
                        if overlap:
                            st.warning(
                                f"**{overlap:,} dari {len(branch_rows):,} baris** sudah ada "
                                f"di database untuk tenant ini pada tanggal dan jam yang sama. "
                                f"Baris tersebut akan **ditimpa**, bukan ditambahkan — "
                                f"jadi total penjualan tidak akan berlipat."
                            )

                        preview = pd.DataFrame(branch_rows[:5])
                        st.dataframe(preview, use_container_width=True, hide_index=True)

                        if st.button(f"✅ Upload {len(insert_rows)} baris", key=f"u_{branch}_{file.name}"):
                            with st.spinner("Menyimpan ke database..."):
                                db.append_sales_data(insert_rows)
                                added = len(insert_rows)
                                db.log_upload(tenant_name, file.name, added, user["email"])
                            st.success(f"✅ **{added}** baris berhasil diupload!")
                            if added > 0:
                                st.balloons()

                except Exception as e:
                    st.error(f"❌ Error parsing: {e}")


# ============================================================================
# TENANT MANAGEMENT
# ============================================================================
def _unified_sales(db):
    """
    F&B and Playground on one axis: tenant, date, revenue, visitors.

    The two POS systems record different things -- F&B counts pax per hour,
    Playground counts children and companions per transaction -- so they are
    reduced to the two measures that mean the same thing on both sides. That
    is what makes a single property-wide leaderboard honest.
    """
    parts = []
    fnb = db.get_sales_data()
    if not fnb.empty:
        parts.append(pd.DataFrame({
            "tenant_id":   fnb["tenant_id"],
            "tenant_name": fnb["tenant_name"],
            "category":    fnb["category"] if "category" in fnb.columns else "F&B",
            "unit_code":   fnb["unit_code"] if "unit_code" in fnb.columns else None,
            "sales_date":  fnb["sales_date"],
            "nett_sales":  fnb["nett_sales"],
            "visitors":    fnb["pax_total"],
        }))
    pg = db.get_playground_data()
    if not pg.empty:
        pg_id, name = db.get_playground_tenant_id(), "Playground"
        tenants = db.get_tenants()
        if pg_id is not None and not tenants.empty:
            m = tenants[tenants["tenant_id"] == pg_id]
            if not m.empty: name = m.iloc[0]["tenant_name"]
        parts.append(pd.DataFrame({
            "tenant_id":   pg_id,
            "tenant_name": name,
            "category":    "Playground",
            "unit_code":   None,
            "sales_date":  pg["sales_date"],
            "nett_sales":  pg["nett_sales"],
            "visitors":    pg["child_total"] + pg["companion_total"],
        }))
    if not parts: return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    out["month"] = out["sales_date"].dt.to_period("M").astype(str)
    return out


def page_performance():
    render_header()
    db = get_db()
    df = _unified_sales(db)

    if df.empty:
        render_page_head("Performa Tenant", "Belum ada data penjualan")
        render_card("Belum ada data",
                    '<p style="font-size:12.5px;color:#8B948D;margin:0">Upload file ESB '
                    'atau CSV Playground lebih dulu — halaman ini akan terisi otomatis.</p>')
        return

    months = sorted(df["month"].unique())
    # Selector sits beside the title, not stacked above it on bare canvas.
    # Streamlit fills a column wherever the call happens, so the period can be
    # read first and the subtitle that depends on it written afterwards.
    hc = st.columns([3, 1])
    with hc[1]:
        month = st.selectbox("Periode", months, index=len(months) - 1,
                             label_visibility="collapsed")
    prev = months[months.index(month) - 1] if months.index(month) > 0 else None

    cur_df = df[df["month"] == month]
    prev_df = df[df["month"] == prev] if prev else df.iloc[0:0]

    # A month still in progress must never be measured against a complete one.
    cutoff = None
    last_day = cur_df["sales_date"].max()
    if pd.notna(last_day) and last_day.date() < pd.Period(month, "M").end_time.date():
        cutoff = int(last_day.day)
        if prev:
            prev_df = prev_df[prev_df["sales_date"].dt.day <= cutoff]

    if not prev:
        sub = f"{month} · tidak ada periode sebelumnya untuk dibandingkan"
    elif cutoff:
        sub = f"{month} · vs {prev}, dibatasi tanggal 1–{cutoff} di kedua sisi"
    else:
        sub = f"{month} · vs {prev}, bulan penuh"
    with hc[0]:
        render_page_head("Performa Tenant", sub)

    def _agg(d):
        return (d.groupby(["tenant_id", "tenant_name"], dropna=False)
                 .agg(sales=("nett_sales", "sum"), visitors=("visitors", "sum"))
                 .reset_index())

    cur, before = _agg(cur_df), _agg(prev_df)
    total, total_prev = cur["sales"].sum(), before["sales"].sum()
    vis, vis_prev = cur["visitors"].sum(), before["visitors"].sum()
    check = total / vis if vis else 0
    check_prev = total_prev / vis_prev if vis_prev else 0
    pct = lambda a, b: ((a - b) / b * 100) if b else None
    id_fmt = lambda n: f"{n:,.0f}".replace(",", ".")

    prev_map = dict(zip(before["tenant_id"], before["sales"]))
    board = cur.copy()
    board["mom"] = board.apply(
        lambda r: pct(r["sales"], prev_map.get(r["tenant_id"], 0)), axis=1)
    board["avg_check"] = board.apply(
        lambda r: round(r["sales"] / r["visitors"]) if r["visitors"] else 0, axis=1)
    board = board.sort_values("sales", ascending=False)
    falling = board[board["mom"].notna() & (board["mom"] < -10)].sort_values("mom")

    # ---------------- KPI ----------------
    k = st.columns(4)
    with k[0]:
        render_kpi("Nett Sales", fmt_rp(total), pct(total, total_prev), featured=True,
                   caption=f"dari {fmt_rp(total_prev)}" if total_prev else None)
    with k[1]:
        render_kpi("Pengunjung", id_fmt(vis), pct(vis, vis_prev),
                   caption=f"dari {id_fmt(vis_prev)}" if vis_prev else None)
    with k[2]:
        render_kpi("Rata / Pengunjung", fmt_rp(check), pct(check, check_prev),
                   caption=f"dari {fmt_rp(check_prev)}" if check_prev else None)
    with k[3]:
        render_kpi("Perlu Perhatian", f"{len(falling)}",
                   caption="tenant turun lebih dari 10%")

    # ---------------- chart | alert | tenant ----------------
    c = st.columns([2, 1, 1])
    with c[0]:
        # Holds a chart, so it is a bordered container rather than HTML.
        with st.container(border=True):
            st.markdown("###### Penjualan per hari")
            days = ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"]
            daily = (cur_df.assign(dow=cur_df["sales_date"].dt.dayofweek)
                           .groupby("dow")["nett_sales"].sum()
                           .reindex(range(7), fill_value=0))
            vals = [float(daily.get(d, 0)) for d in [6, 0, 1, 2, 3, 4, 5]]
            active = [v for v in vals if v > 0]
            avg = sum(active) / len(active) if active else 0
            # Hatched below the weekly average, solid at or above it. Days with
            # no trading at all are left flat rather than drawn as a stub,
            # which would read as a real but tiny figure.
            fig = go.Figure(go.Bar(
                x=days, y=vals,
                marker=dict(
                    color=["#1A6B3F" if v >= avg and v > 0 else "#F3F4F1" for v in vals],
                    cornerradius=18,
                    pattern=dict(shape=["" if (v >= avg and v > 0) else "/" for v in vals],
                                 fgcolor="#D9DDD7", size=5, solidity=.3)),
                hovertemplate="%{x}<br>Rp %{y:,.0f}<extra></extra>"))
            fig.update_layout(height=200, margin=dict(t=4, b=4, l=4, r=4),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              yaxis=dict(visible=False), showlegend=False, bargap=.4,
                              xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#8B948D")))
            show_chart(fig, use_container_width=True, config={"displayModeBar": False})
            st.caption("Bar pekat menandai hari di atas rata-rata, "
                       "bar diarsir di bawahnya.")

    with c[1]:
        if not falling.empty:
            w = falling.iloc[0]
            render_card(
                "Perlu perhatian",
                f'<p class="big-alert">{w["tenant_name"]} turun {abs(w["mom"]):.1f}%</p>'
                f'<p style="font-size:11.5px;color:#8B948D;line-height:1.45;margin:0">'
                f'{fmt_rp(prev_map.get(w["tenant_id"], 0))} &rarr; {fmt_rp(w["sales"])} '
                f'pada periode ini.</p>',
                foot=f"{len(falling)} tenant turun lebih dari 10%.")
        else:
            render_card(
                "Perlu perhatian",
                '<p class="big-alert" style="color:#1A6B3F">Tidak ada</p>'
                '<p style="font-size:11.5px;color:#8B948D;margin:0">Tidak ada tenant yang '
                'turun lebih dari 10% dibanding periode sebelumnya.</p>')

    with c[2]:
        render_card("Tenant",
                    rows_html([(r["tenant_id"], r["tenant_name"],
                                f'{id_fmt(r["visitors"])} pengunjung',
                                f'<span class="amt num">'
                                f'{fmt_rp(r["sales"]).replace("Rp ", "")}</span>')
                               for _, r in board.iterrows()]),
                    foot="Warna tiap tenant tetap sama di seluruh aplikasi.")

    # ---------------- movement | target | occupancy ----------------
    c2 = st.columns([2, 1, 1])
    with c2[0]:
        render_card("Pergerakan bulan ini",
                    rows_html([(r["tenant_id"], r["tenant_name"],
                                f'{fmt_rp(r["avg_check"])} per pengunjung · '
                                f'{id_fmt(r["visitors"])} orang',
                                move_tag(r["mom"]))
                               for _, r in board.iterrows()]))

    with c2[1]:
        targets = db.get_targets()
        tmap = {}
        if not targets.empty:
            sel = targets[targets["period_month"].astype(str).str[:7] == month]
            tmap = dict(zip(sel["tenant_id"], sel["target_nett"])) if not sel.empty else {}
        target_total = sum(tmap.get(t, 0) for t in board["tenant_id"])
        if target_total:
            render_card("Capaian target",
                        gauge_html(total / target_total, f"dari {fmt_rp(target_total)}"))
        else:
            render_card("Capaian target", gauge_html(0, "target belum diisi"),
                        foot="Isi target di <b>Kelola Unit &amp; Sewa &rarr; Target</b> "
                             "untuk mengaktifkan kartu ini.")

    with c2[2]:
        units, tenancies = db.get_units(), db.get_tenancies()
        n_units = len(units) if not units.empty else 0
        occupied = 0
        if not tenancies.empty and n_units:
            occupied = tenancies[tenancies["end_date"].isna()]["unit_id"].nunique()
        bars = "".join(
            f'<i style="flex:1;height:5px;border-radius:20px;background:'
            f'{"#2E9160" if i < occupied else "rgba(255,255,255,.2)"}"></i>'
            for i in range(max(n_units, 1)))
        empty = n_units - occupied
        with st.container(border=True):
            st.metric("Okupansi unit", f"{occupied} / {n_units}")
            st.progress(occupied / n_units if n_units else 0.0)
            st.caption("Seluruh unit terisi" if empty == 0 and n_units
                       else f"{empty} unit tanpa penyewa aktif")

    # ---------------- per unit ----------------
    if "unit_code" in cur_df.columns and cur_df["unit_code"].notna().any():
        per_unit = (cur_df.dropna(subset=["unit_code"]).groupby("unit_code")
                          .agg(sales=("nett_sales", "sum"), visitors=("visitors", "sum"),
                               brand=("tenant_name", lambda s: ", ".join(sorted(set(s)))))
                          .reset_index().sort_values("sales", ascending=False))
        with st.container(border=True):
            st.markdown("###### Performa per unit")
            st.dataframe(per_unit, use_container_width=True, hide_index=True,
                column_config={
                    "unit_code": st.column_config.TextColumn("Unit"),
                    "brand":     st.column_config.TextColumn("Brand pada periode ini"),
                    "sales":     st.column_config.NumberColumn("Nett Sales (Rp)", format="localized"),
                    "visitors":  st.column_config.NumberColumn("Pengunjung", format="localized"),
                })
            st.caption("Unit yang berganti penyewa di tengah periode menjumlahkan "
                       "kontribusi kedua brand sesuai tanggalnya.")

def _occupancy_table(units, tenancies):
    """One row per unit, with whoever occupies it today."""
    rows = []
    for _, u in units.iterrows():
        active = pd.DataFrame()
        if not tenancies.empty:
            active = tenancies[(tenancies["unit_id"] == u["unit_id"]) &
                               (tenancies["end_date"].isna())]
        if not active.empty:
            a = active.iloc[0]
            rows.append({"Unit": u["unit_code"], "Penyewa": a.get("tenant_name", "—"),
                         "Sejak": str(a["start_date"])[:10], "Status": "🟢 Terisi"})
        else:
            rows.append({"Unit": u["unit_code"], "Penyewa": "—",
                         "Sejak": "—", "Status": "⚪ Kosong"})
    return pd.DataFrame(rows)


def page_units():
    st.markdown("## 🏬 Kelola Unit & Masa Sewa")
    st.caption(
        "Unit adalah ruang fisiknya dan sifatnya tetap; tenant adalah brand yang "
        "menempatinya dan bisa berganti. Penjualan selalu melekat pada brand, "
        "sehingga pergantian penyewa tidak pernah mengubah sejarah siapa pun."
    )
    db = get_db()
    units, tenants = db.get_units(), db.get_tenants()
    tenancies = db.get_tenancies_detailed()
    user = st.session_state["user"]

    t_occ, t_swap, t_unit, t_hist, t_target = st.tabs([
        "🏢 Okupansi", "🔄 Ganti Penyewa", "➕ Unit", "🕓 Riwayat Sewa", "🎯 Target"])

    # ---------- Okupansi saat ini ----------
    with t_occ:
        if units.empty:
            st.info("Belum ada unit. Tambahkan di tab **➕ Unit**.")
        else:
            occ = _occupancy_table(units, tenancies)
            terisi = int((occ["Status"] == "🟢 Terisi").sum())
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Unit", len(occ))
            c2.metric("Terisi", terisi)
            c3.metric("Kosong", len(occ) - terisi,
                      delta=None if len(occ) == terisi else "perlu tenant")
            st.dataframe(occ, use_container_width=True, hide_index=True)

    # ---------- Ganti penyewa ----------
    with t_swap:
        if units.empty or tenants.empty:
            st.info("Perlu minimal satu unit dan satu tenant terdaftar.")
        else:
            st.markdown("#### Serah terima unit ke brand lain")
            unit_map = {f"{r['unit_code']} ({r['unit_id']})": r for _, r in units.iterrows()}
            pick = st.selectbox("Unit", list(unit_map.keys()))
            unit = unit_map[pick]
            current = db.get_active_tenancy(unit["unit_id"])

            if current:
                cur_name = tenants[tenants["tenant_id"] == current["tenant_id"]]
                cur_name = cur_name.iloc[0]["tenant_name"] if not cur_name.empty else current["tenant_id"]
                st.info(f"Penyewa saat ini: **{cur_name}** — sejak {str(current['start_date'])[:10]}")
            else:
                st.warning("Unit ini sedang kosong. Pengisian akan dicatat sebagai penyewa pertama.")

            avail = tenants if current is None else tenants[tenants["tenant_id"] != current["tenant_id"]]
            if avail.empty:
                st.error("Tidak ada brand lain yang bisa dipilih. Daftarkan dulu di **Kelola Tenant**.")
            else:
                tmap = {f"{r['tenant_name']} ({r['tenant_id']})": r for _, r in avail.iterrows()}
                new_pick = st.selectbox("Penyewa baru", list(tmap.keys()))
                new_tenant = tmap[new_pick]
                handover = st.date_input("Tanggal mulai penyewa baru", value=date.today())

                st.markdown(
                    f"Penjualan **sampai {handover - timedelta(days=1):%d %b %Y}** tetap "
                    f"tercatat atas penyewa lama, dan mulai **{handover:%d %b %Y}** "
                    f"tercatat atas **{new_tenant['tenant_name']}**. "
                    "Tidak ada satu baris penjualan pun yang diubah."
                )
                if st.button("🔄 Proses serah terima", type="primary", use_container_width=True):
                    try:
                        db.replace_tenant(unit["unit_id"], new_tenant["tenant_id"], handover)
                        st.success(
                            f"✅ **{unit['unit_code']}** kini ditempati "
                            f"**{new_tenant['tenant_name']}** sejak {handover:%d %b %Y}."
                        )
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(f"Gagal memproses: {e}")

    # ---------- Kelola unit ----------
    with t_unit:
        st.markdown("#### Tambah unit")
        with st.form("add_unit"):
            c1, c2 = st.columns(2)
            with c1:
                u_code = st.text_input("Kode Unit", placeholder="mis. Slot DOD, A-3")
                u_floor = st.text_input("Lantai", placeholder="mis. GF")
            with c2:
                u_name = st.text_input("Keterangan", placeholder="opsional")
                u_area = st.number_input("Luas (m²)", min_value=0.0, step=0.5, value=0.0)
            if st.form_submit_button("➕ Tambah Unit", use_container_width=True):
                if not u_code:
                    st.error("Kode unit wajib diisi.")
                else:
                    db.add_unit(u_code, u_name or None, u_floor or None, u_area or None)
                    st.success(f"✅ Unit **{u_code}** ditambahkan.")
                    st.rerun()

        if not units.empty:
            st.markdown("#### Ubah unit yang ada")
            st.caption("Edit langsung di tabel, lalu tekan Simpan.")
            editable = units[["unit_id", "unit_code", "unit_name", "floor", "area_sqm", "is_active"]].copy()
            edited = st.data_editor(
                editable, hide_index=True, use_container_width=True, key="unit_editor",
                disabled=["unit_id"],
                column_config={
                    "unit_id":   st.column_config.TextColumn("ID", width="small"),
                    "unit_code": st.column_config.TextColumn("Kode Unit", required=True),
                    "unit_name": st.column_config.TextColumn("Keterangan"),
                    "floor":     st.column_config.TextColumn("Lantai", width="small"),
                    "area_sqm":  st.column_config.NumberColumn("Luas (m²)", format="%.1f"),
                    "is_active": st.column_config.CheckboxColumn("Aktif"),
                })
            if st.button("💾 Simpan perubahan unit", use_container_width=True):
                changed = 0
                for _, row in edited.iterrows():
                    before = editable[editable["unit_id"] == row["unit_id"]].iloc[0]
                    if not before.equals(row):
                        db.edit_unit(row["unit_id"], unit_code=row["unit_code"],
                                     unit_name=row["unit_name"], floor=row["floor"],
                                     area_sqm=float(row["area_sqm"]) if pd.notna(row["area_sqm"]) else None,
                                     is_active=bool(row["is_active"]))
                        changed += 1
                st.success(f"✅ {changed} unit diperbarui." if changed else "Tidak ada perubahan.")
                if changed: st.rerun()

    # ---------- Riwayat sewa ----------
    with t_hist:
        if tenancies.empty:
            st.info("Belum ada riwayat sewa.")
        else:
            show = tenancies.copy()
            show["end_date"] = show["end_date"].fillna("— sekarang")
            cols = [c for c in ["unit_code", "tenant_name", "start_date", "end_date", "status", "notes"]
                    if c in show.columns]
            st.dataframe(
                show[cols], use_container_width=True, hide_index=True,
                column_config={
                    "unit_code":   "Unit",
                    "tenant_name": "Penyewa",
                    "start_date":  "Mulai",
                    "end_date":    "Selesai",
                    "status":      "Status",
                    "notes":       "Catatan",
                })
            st.caption(
                "Baris berstatus **Aktif** adalah penyewa yang sedang menempati. "
                "Database menolak dua masa sewa yang tumpang tindih pada unit yang sama, "
                "karena tumpang tindih akan menggandakan pendapatan unit itu di setiap laporan."
            )

    # ---------- Target ----------
    with t_target:
        if tenants.empty:
            st.info("Belum ada tenant.")
        else:
            st.markdown("#### Target penjualan bulanan")
            period = st.date_input("Bulan", value=date.today().replace(day=1),
                                   help="Tanggal berapa pun; yang dipakai bulannya.")
            month_start = period.replace(day=1)

            existing = db.get_targets()
            cur = {}
            if not existing.empty:
                sel = existing[existing["period_month"].astype(str).str[:7] == f"{month_start:%Y-%m}"]
                cur = dict(zip(sel["tenant_id"], sel["target_nett"])) if not sel.empty else {}

            tgt_df = pd.DataFrame({
                "tenant_id":   tenants["tenant_id"],
                "Tenant":      tenants["tenant_name"],
                "Target (Rp)": [int(cur.get(t, 0)) for t in tenants["tenant_id"]],
            })
            edited_t = st.data_editor(
                tgt_df, hide_index=True, use_container_width=True, key="target_editor",
                disabled=["tenant_id", "Tenant"],
                column_config={
                    "tenant_id":   None,
                    "Tenant":      st.column_config.TextColumn("Tenant"),
                    "Target (Rp)": st.column_config.NumberColumn("Target (Rp)", min_value=0, step=1_000_000, format="%d"),
                })
            if st.button(f"💾 Simpan target {month_start:%B %Y}", use_container_width=True):
                n = 0
                for _, r in edited_t.iterrows():
                    db.upsert_target(r["tenant_id"], month_start, r["Target (Rp)"], user["email"])
                    n += 1
                st.success(f"✅ Target {month_start:%B %Y} tersimpan untuk {n} tenant.")
                st.rerun()


def page_tenants():
    render_header()
    db = get_db()
    st.subheader("🏢 Kelola Tenant")
    tdf = db.get_tenants()
    if not tdf.empty:
        st.dataframe(tdf, use_container_width=True, hide_index=True)

    st.divider()

    tab_add, tab_edit, tab_delete = st.tabs(["➕ Tambah Tenant", "✏️ Edit Tenant", "🗑️ Hapus Tenant"])

    # --- TAB: Tambah ---
    with tab_add:
        with st.form("add_tenant"):
            c1,c2 = st.columns(2)
            with c1: t_name = st.text_input("Nama Tenant", placeholder="e.g. Kopi Kenangan CIBIS")
            with c2: t_branch = st.text_input("ESB Branch Name (persis seperti di file)", placeholder="e.g. KOPI_KENANGAN_CIBIS")
            if st.form_submit_button("➕ Tambah Tenant", use_container_width=True):
                if t_name and t_branch:
                    db.add_tenant(t_name, t_branch)
                    st.success(f"✅ Tenant '{t_name}' berhasil ditambahkan.")
                    st.rerun()
                else:
                    st.error("Semua field wajib diisi.")

    # --- TAB: Edit ---
    with tab_edit:
        if tdf.empty:
            st.info("Belum ada tenant untuk diedit.")
        else:
            tenant_options = {f"{r['tenant_name']} ({r['tenant_id']})": r for _, r in tdf.iterrows()}
            selected = st.selectbox("Pilih Tenant", list(tenant_options.keys()), key="edit_sel")
            tenant = tenant_options[selected]

            with st.form("edit_tenant"):
                c1, c2 = st.columns(2)
                with c1:
                    new_name = st.text_input("Nama Tenant", value=tenant["tenant_name"])
                with c2:
                    new_branch = st.text_input("ESB Branch Name", value=tenant["esb_branch_name"])

                sales_df = db.get_sales_data(tenant["tenant_name"])
                sales_count = len(sales_df) if not sales_df.empty else 0
                st.caption(f"ℹ️ Tenant ini memiliki **{sales_count:,}** baris data sales. "
                           f"Jika nama diubah, seluruh data akan otomatis menyesuaikan.")

                if st.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
                    if not new_name or not new_branch:
                        st.error("Semua field wajib diisi.")
                    else:
                        ok = db.edit_tenant(tenant["tenant_id"], new_name, new_branch)
                        if ok:
                            st.success(f"✅ Tenant berhasil diupdate → **{new_name}** / `{new_branch}`")
                            st.rerun()
                        else:
                            st.error("Gagal mengupdate tenant.")

    # --- TAB: Hapus ---
    with tab_delete:
        if tdf.empty:
            st.info("Belum ada tenant untuk dihapus.")
        else:
            del_options = {f"{r['tenant_name']} ({r['tenant_id']})": r for _, r in tdf.iterrows()}
            del_selected = st.selectbox("Pilih Tenant", list(del_options.keys()), key="del_sel")
            del_tenant = del_options[del_selected]

            sales_df2 = db.get_sales_data(del_tenant["tenant_name"])
            sales_count = len(sales_df2) if not sales_df2.empty else 0

            st.warning(
                f"⚠️ **Menghapus tenant '{del_tenant['tenant_name']}'** akan menghapus:\n"
                f"- Data tenant dari daftar\n"
                f"- **{sales_count:,} baris** data sales terkait\n"
                f"- Seluruh log upload terkait\n\n"
                f"**Aksi ini tidak dapat dibatalkan.**"
            )

            if st.button(f"🗑️ Ya, Hapus '{del_tenant['tenant_name']}'",
                         type="primary", use_container_width=True, key="confirm_del"):
                deleted, sales_del = db.delete_tenant(del_tenant["tenant_id"])
                if deleted:
                    st.success(f"✅ Tenant **{del_tenant['tenant_name']}** dihapus beserta {sales_del:,} baris data sales.")
                    st.rerun()
                else:
                    st.error("Gagal menghapus tenant.")


# ============================================================================
# USER MANAGEMENT
# ============================================================================
def page_users():
    render_header()
    db = get_db()
    st.subheader("👥 Kelola User")
    udf = db.get_all_users()
    if not udf.empty:
        st.dataframe(udf, use_container_width=True, hide_index=True)
    st.divider()
    c1,c2 = st.columns(2)
    with c1:
        st.markdown("**Tambah / Update Role User**")
        st.caption("User dengan domain @srkel.id atau @teamup.id akan otomatis terdaftar sebagai Viewer saat pertama kali login. Gunakan form ini untuk mengubah role.")
        with st.form("add_user"):
            email = st.text_input("Email User", placeholder="nama@srkel.id")
            dn = st.text_input("Nama Lengkap")
            role = st.selectbox("Role", ["Viewer","Admin","Super Admin"])
            tdf = db.get_tenants()
            if not tdf.empty and role == "Viewer":
                acc = st.multiselect("Akses Tenant", tdf["tenant_name"].tolist())
                acc_str = ",".join(acc) if acc else "ALL"
            else:
                acc_str = "ALL"
                if role != "Viewer":
                    st.info("Admin & Super Admin → akses semua tenant.")
            if st.form_submit_button("💾 Simpan", use_container_width=True):
                if email and dn:
                    existing = db.get_user(email)
                    if existing:
                        db.update_user_role(email, role, acc_str)
                        st.success(f"✅ Role '{email}' diupdate ke {role}.")
                    else:
                        db.create_user(email, dn, role, acc_str, st.session_state["user"]["email"])
                        st.success(f"✅ User '{email}' ditambahkan sebagai {role}.")
                    st.rerun()
    with c2:
        st.markdown("**Update Status User**")
        if not udf.empty:
            with st.form("upd_user"):
                target = st.selectbox("Pilih User", udf["email"].tolist())
                new_st = st.selectbox("Status", ["TRUE","FALSE"], format_func=lambda x: "Aktif" if x=="TRUE" else "Nonaktif")
                if st.form_submit_button("🔄 Update", use_container_width=True):
                    db.update_user_status(target, new_st)
                    st.success(f"✅ Status '{target}' diupdate.")
                    st.rerun()


# ============================================================================
# UPLOAD LOG
# ============================================================================
def page_upload_log():
    render_header()
    db = get_db()
    st.subheader("📋 Log Upload")
    log = db.get_upload_log()
    if not log.empty:
        st.dataframe(log, use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada riwayat upload.")


# ============================================================================
# UPLOAD PLAYGROUND
# ============================================================================
def page_upload_playground():
    render_header()
    db = get_db()
    user = st.session_state["user"]

    st.subheader("🎪 Upload Data Playground TnT")
    st.markdown("""
    <div class="upload-info">
        <strong>Petunjuk:</strong> Upload file CSV dari POS Playground
        (<em>pos_sales_report</em>). Sistem akan otomatis mendeteksi
        dan menambahkan data baru berdasarkan Order ID (tanpa duplikasi).<br>
        <strong>💡 Bisa upload beberapa file sekaligus!</strong>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader("Pilih file Playground (.csv)", type=["csv"], accept_multiple_files=True, key="pg_upload")

    if uploaded:
        for file in uploaded:
            with st.expander(f"📄 {file.name}", expanded=True):
                try:
                    raw = file.read()
                    rows = PlaygroundParser.parse(raw, file.name)
                    if not rows:
                        st.error("File tidak berisi data valid.")
                        continue

                    st.markdown(f"**Data ditemukan:** {len(rows)} transaksi")
                    date_min = min(r["sales_date"] for r in rows)
                    date_max = max(r["sales_date"] for r in rows)
                    total_amount = sum(r["amount"] for r in rows)
                    total_child = sum(r["child_total"] for r in rows)
                    st.markdown(f"**Periode:** {date_min} s.d. {date_max} · **Total:** Rp {total_amount:,} · **Children:** {total_child:,}")

                    preview = pd.DataFrame(rows[:5])
                    st.dataframe(preview, use_container_width=True, hide_index=True)

                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    insert_rows = [(r["order_id"],r["amount"],r["nett_sales"],r["tax_amount"],
                                    r["customer_name"],r["sales_date"],r["child_total"],
                                    r["companion_total"],user["email"],now) for r in rows]

                    if st.button(f"✅ Upload {len(insert_rows)} transaksi", key=f"pg_{file.name}"):
                        with st.spinner("Menyimpan ke database..."):
                            added = db.insert_playground_batch(insert_rows)
                            db.log_upload("Playground TnT", file.name, added, user["email"])
                        skipped = len(insert_rows) - added
                        st.success(f"✅ **{added}** transaksi baru" + (f" · {skipped} duplikat di-skip" if skipped else ""))
                        if added > 0: st.balloons()
                except Exception as e:
                    st.error(f"❌ Error parsing: {e}")


# ============================================================================
# DASHBOARD PLAYGROUND
# ============================================================================
def _export_row(df, label, dr, kind="fnb"):
    """The four export buttons, identical on both dashboards."""
    s, e = (dr[0], dr[1]) if isinstance(dr, tuple) else (date.today(), date.today())
    with st.container(border=True):
        st.markdown("###### Ekspor")
        c = st.columns(4)
        if kind == "fnb":
            with c[0]:
                st.download_button("🖨️ Print / PDF", generate_dashboard_html(df, label, s, e),
                                   f"grove_fnb_{s}_{e}.html", "text/html", use_container_width=True)
            with c[1]:
                st.download_button("📊 XLSX", generate_dashboard_xlsx(df, label, s, e),
                                   f"grove_fnb_{s}_{e}.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        else:
            with c[0]:
                st.download_button("🖨️ Print / PDF", generate_playground_html(df.copy(), s, e),
                                   f"grove_playground_{s}_{e}.html", "text/html",
                                   use_container_width=True)
            with c[1]:
                st.download_button("📊 XLSX", generate_playground_xlsx(df.copy(), s, e),
                                   f"grove_playground_{s}_{e}.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        with c[2]:
            st.download_button("⬇️ Data mentah (CSV)", df.to_csv(index=False).encode(),
                               f"grove_{kind}_raw.csv", "text/csv", use_container_width=True)
        with c[3]:
            g = df.groupby(df["sales_date"].dt.date)["nett_sales"].sum().reset_index()
            st.download_button("⬇️ Ringkasan (CSV)", g.to_csv(index=False).encode(),
                               f"grove_{kind}_summary.csv", "text/csv", use_container_width=True)


def page_dashboard_playground():
    render_header()
    db = get_db()
    df_all = db.get_playground_data()

    hc = st.columns([3, 1])
    if df_all.empty:
        with hc[0]:
            render_page_head("Dashboard Playground", "Belum ada data")
        render_card("Belum ada data",
                    '<p style="font-size:12.5px;color:#8B948D;margin:0">Upload file CSV '
                    'melalui menu Upload Playground.</p>')
        return
    # Children and companions are the two things being counted; visitors is the
    # sum, which is what makes Playground comparable with an F&B pax count.
    df_all = df_all.assign(visitors=df_all["child_total"] + df_all["companion_total"])
    with hc[1]:
        df, dr = _filters(df_all, "pg_dr")
    with hc[0]:
        render_page_head("Dashboard Playground",
                         f"Twist N' Turns · {dr[0]:%d %b %Y} – {dr[1]:%d %b %Y}")
    if df.empty:
        render_card("Tidak ada data",
                    '<p style="font-size:12.5px;color:#8B948D;margin:0">Rentang tanggal '
                    'yang dipilih tidak memuat transaksi.</p>')
        return

    cfg = dict(val="nett_sales", cnt="visitors", cnt_label="Pengunjung",
               per_label="per pengunjung")
    nett = df["nett_sales"].sum()
    trx, child, comp = len(df), df["child_total"].sum(), df["companion_total"].sum()
    ndays = df["sales_date"].dt.date.nunique()

    k = st.columns(4)
    with k[0]: render_kpi("Nett Sales", fmt_rp(nett), featured=True,
                          caption=f"{ndays} hari beroperasi")
    with k[1]: render_kpi("Transaksi", f"{trx:,.0f}".replace(",", "."),
                          caption=f"{trx/ndays:,.1f} per hari".replace(",", ".") if ndays else None)
    with k[2]: render_kpi("Anak", f"{child:,.0f}".replace(",", "."),
                          caption=f"{child/trx:.2f} per transaksi" if trx else None)
    with k[3]: render_kpi("Rata per Transaksi", fmt_rp(nett / trx if trx else 0),
                          caption="nett dibagi jumlah transaksi")

    t = st.tabs(["Trend Harian", "Child & Companion", "Weekend vs Weekday",
                 "Weekly Report", "Monthly Overview", "Deep Dive"])

    with t[0]:
        sec_trend(df, cfg)

    with t[1]:
        daily = (df.groupby(df["sales_date"].dt.date)
                   .agg(anak=("child_total", "sum"), pendamping=("companion_total", "sum"))
                   .reset_index())
        daily.columns = ["day", "anak", "pendamping"]
        c = st.columns([2, 1])
        with c[0]:
            with st.container(border=True):
                st.markdown("###### Anak dan pendamping per hari")
                fig = go.Figure()
                fig.add_bar(x=daily["day"], y=daily["anak"], name="Anak",
                            marker_color=GREEN,
                            hovertemplate="%{x|%d %b}<br>%{y:,.0f} anak<extra></extra>")
                fig.add_bar(x=daily["day"], y=daily["pendamping"], name="Pendamping",
                            marker_color=OCHRE,
                            hovertemplate="%{x|%d %b}<br>%{y:,.0f} pendamping<extra></extra>")
                fig.update_layout(barmode="stack", height=250,
                                  margin=dict(t=10, b=34, l=54, r=14))
                show_chart(fig)
        with c[1]:
            ratio = comp / child if child else 0
            solo = int((df["companion_total"] == 0).sum())
            render_card("Komposisi kunjungan", pd.DataFrame([
                {"Ukuran": "Rasio pendamping", "Keterangan": "pendamping per satu anak",
                 "Nilai": f"{ratio:.2f}"},
                {"Ukuran": "Total pendamping", "Keterangan": f"{comp/trx:.2f} per transaksi",
                 "Nilai": id_num(comp)},
                {"Ukuran": "Tanpa pendamping",
                 "Keterangan": f"{solo/trx*100:.1f}% dari transaksi",
                 "Nilai": id_num(solo)}]),
                foot="Rasio di bawah 1 berarti sebagian anak datang tanpa pendamping berbayar.")

    with t[2]:
        we = df.assign(we=df["sales_date"].dt.dayofweek >= 5)
        agg = we.groupby("we").agg(v=("nett_sales", "sum"), t=("order_id", "count"),
                                   a=("child_total", "sum"),
                                   d=("sales_date", lambda s: s.dt.date.nunique())).reset_index()
        agg["label"] = agg["we"].map({False: "Hari kerja", True: "Akhir pekan"})
        agg["per_day"] = agg["v"] / agg["d"].replace(0, pd.NA)
        c = st.columns([1, 2])
        with c[0]:
            with st.container(border=True):
                st.markdown("###### Rata-rata per hari")
                fig = go.Figure(go.Bar(
                    x=agg["label"], y=agg["per_day"],
                    marker=dict(color=[GREEN, OCHRE][:len(agg)]),
                    hovertemplate="%{x}<br>Rp %{y:,.0f} per hari<extra></extra>"))
                fig.update_layout(height=240, showlegend=False, bargap=.45,
                                  yaxis=dict(visible=False), margin=dict(t=10, b=34, l=10, r=10))
                show_chart(fig)
        with c[1]:
            render_card("Hari kerja dibanding akhir pekan",
                rows_html([(r["label"], r["label"],
                            f"{r['d']} hari · {r['t']:,.0f} transaksi · {r['a']:,.0f} anak"
                            .replace(",", "."),
                            f'<span class="amt num">{fmt_rp(r["per_day"])} / hari</span>')
                           for _, r in agg.iterrows()]),
                foot="Dibandingkan sebagai rata-rata harian, karena akhir pekan hanya "
                     "dua dari tujuh hari — total mentahnya akan selalu kalah.")

    with t[3]: sec_weekly(df, cfg)
    with t[4]: sec_monthly(df, cfg)
    with t[5]: sec_deep(df, cfg)

    _export_row(df, "Playground", dr, kind="pg")


def page_master_dashboard():
    render_header()
    db = get_db()
    df_fnb, df_pg = db.get_sales_data(), db.get_playground_data()
    fnb_empty, pg_empty = df_fnb.empty, df_pg.empty

    if fnb_empty and pg_empty:
        render_page_head("Master Dashboard", "Belum ada data")
        render_card("Belum ada data",
                    '<p style="font-size:12.5px;color:#8B948D;margin:0">Upload data F&amp;B '
                    'atau Playground lebih dulu.</p>')
        return

    dates = []
    if not fnb_empty: dates += df_fnb["sales_date"].tolist()
    if not pg_empty:  dates += df_pg["sales_date"].tolist()
    dmin, dmax = min(dates).date(), max(dates).date()

    hc = st.columns([3, 1])
    with hc[1]:
        dr = st.date_input("Rentang tanggal",
                           value=(max(dmin, dmax - timedelta(days=30)), dmax),
                           min_value=dmin, max_value=dmax, key="master_dr",
                           label_visibility="collapsed")
    if isinstance(dr, tuple) and len(dr) == 2:
        s, e = pd.Timestamp(dr[0]), pd.Timestamp(dr[1])
        if not fnb_empty: df_fnb = df_fnb[(df_fnb["sales_date"] >= s) & (df_fnb["sales_date"] <= e)]
        if not pg_empty:  df_pg = df_pg[(df_pg["sales_date"] >= s) & (df_pg["sales_date"] <= e)]
    else:
        dr = (dmin, dmax)
    fnb_empty, pg_empty = df_fnb.empty, df_pg.empty
    with hc[0]:
        render_page_head("Master Dashboard",
                         f"Seluruh tenant · {dr[0]:%d %b %Y} – {dr[1]:%d %b %Y}")

    fnb_nett = df_fnb["nett_sales"].sum() if not fnb_empty else 0
    fnb_pax = df_fnb["pax_total"].sum() if not fnb_empty else 0
    pg_nett = df_pg["nett_sales"].sum() if not pg_empty else 0
    pg_vis = (df_pg["child_total"] + df_pg["companion_total"]).sum() if not pg_empty else 0
    grand, visitors = fnb_nett + pg_nett, fnb_pax + pg_vis
    idf = lambda n: f"{n:,.0f}".replace(",", ".")

    k = st.columns(4)
    with k[0]: render_kpi("Total Nett Sales", fmt_rp(grand), featured=True,
                          caption="F&B dan Playground digabung")
    with k[1]: render_kpi("F&B", fmt_rp(fnb_nett),
                          caption=f"{fnb_nett/grand*100:.1f}% dari total" if grand else "belum ada data")
    with k[2]: render_kpi("Playground", fmt_rp(pg_nett),
                          caption=f"{pg_nett/grand*100:.1f}% dari total" if grand else None)
    with k[3]: render_kpi("Pengunjung", idf(visitors),
                          caption=f"{idf(fnb_pax)} pax · {idf(pg_vis)} playground")

    # F&B and Playground must be told apart at a glance, so they hold two
    # distinct hues from the palette everywhere on this page. A bulk colour
    # migration had collapsed both onto the same green, which left the
    # comparison charts unable to say which half was which.
    C_FNB, C_PG = GREEN, OCHRE

    t = st.tabs(["Overview", "Perbandingan Harian", "Perbandingan Bulanan"])

    with t[0]:
        c = st.columns([1, 2])
        with c[0]:
            with st.container(border=True):
                st.markdown("###### Kontribusi pendapatan")
                fig = go.Figure(go.Pie(
                    labels=["F&B", "Playground"], values=[fnb_nett, pg_nett], hole=.62,
                    marker=dict(colors=[C_FNB, C_PG], line=dict(color="#fff", width=2)),
                    textinfo="percent",
                    hovertemplate="%{label}<br>Rp %{value:,.0f}<extra></extra>"))
                fig.update_layout(height=260, margin=dict(t=10, b=10, l=10, r=10),
                                  legend=dict(orientation="h", y=-0.05, x=.5, xanchor="center"))
                show_chart(fig)
        with c[1]:
            rows = []
            if not fnb_empty:
                for _, r in (df_fnb.groupby(["tenant_id", "tenant_name"])["nett_sales"]
                                   .sum().reset_index().sort_values("nett_sales", ascending=False)
                                   ).iterrows():
                    rows.append((r["tenant_id"], r["tenant_name"], "F&B",
                                 f'<span class="amt num">{fmt_rp(r["nett_sales"])}</span>'))
            if not pg_empty:
                rows.append(("T900", "Twist N' Turns", "Playground",
                             f'<span class="amt num">{fmt_rp(pg_nett)}</span>'))
            render_card("Pendapatan per unit usaha", rows_html(rows),
                        foot="Playground memakai POS terpisah, jadi tampil sebagai satu baris.")

    with t[1]:
        d_fnb = (df_fnb.groupby(df_fnb["sales_date"].dt.date)["nett_sales"].sum()
                 if not fnb_empty else pd.Series(dtype=float))
        d_pg = (df_pg.groupby(df_pg["sales_date"].dt.date)["nett_sales"].sum()
                if not pg_empty else pd.Series(dtype=float))
        idx = sorted(set(d_fnb.index) | set(d_pg.index))
        if not idx:
            st.info("Tidak ada data pada rentang ini.")
        else:
            f_v = [float(d_fnb.get(i, 0)) for i in idx]
            p_v = [float(d_pg.get(i, 0)) for i in idx]
            with st.container(border=True):
                st.markdown("###### Pendapatan harian")
                fig = go.Figure()
                fig.add_bar(x=idx, y=f_v, name="F&B", marker_color=C_FNB,
                            hovertemplate="%{x|%d %b}<br>F&B Rp %{y:,.0f}<extra></extra>")
                fig.add_bar(x=idx, y=p_v, name="Playground", marker_color=C_PG,
                            hovertemplate="%{x|%d %b}<br>Playground Rp %{y:,.0f}<extra></extra>")
                fig.update_layout(barmode="stack", height=300,
                                  margin=dict(t=10, b=34, l=56, r=14))
                show_chart(fig)
            with st.container(border=True):
                st.markdown("###### Tren gabungan")
                fig = go.Figure()
                fig.add_scatter(x=idx, y=[a + b for a, b in zip(f_v, p_v)], name="Total",
                                mode="lines", line=dict(color=GREEN_D, width=2.5),
                                fill="tozeroy", fillcolor="rgba(16,61,40,.08)",
                                hovertemplate="%{x|%d %b}<br>Rp %{y:,.0f}<extra></extra>")
                fig.update_layout(height=260, margin=dict(t=10, b=34, l=56, r=14),
                                  showlegend=False)
                show_chart(fig)

    with t[2]:
        m_fnb = (df_fnb.groupby(df_fnb["sales_date"].dt.to_period("M").astype(str))["nett_sales"].sum()
                 if not fnb_empty else pd.Series(dtype=float))
        m_pg = (df_pg.groupby(df_pg["sales_date"].dt.to_period("M").astype(str))["nett_sales"].sum()
                if not pg_empty else pd.Series(dtype=float))
        months = sorted(set(m_fnb.index) | set(m_pg.index))
        if not months:
            st.info("Tidak ada data pada rentang ini.")
        else:
            f_v = [float(m_fnb.get(m, 0)) for m in months]
            p_v = [float(m_pg.get(m, 0)) for m in months]
            with st.container(border=True):
                st.markdown("###### Pendapatan per bulan")
                fig = go.Figure()
                fig.add_bar(x=months, y=f_v, name="F&B", marker_color=C_FNB,
                            hovertemplate="%{x}<br>F&B Rp %{y:,.0f}<extra></extra>")
                fig.add_bar(x=months, y=p_v, name="Playground", marker_color=C_PG,
                            hovertemplate="%{x}<br>Playground Rp %{y:,.0f}<extra></extra>")
                fig.update_layout(barmode="group", height=300, bargap=.3,
                                  margin=dict(t=10, b=34, l=56, r=14))
                show_chart(fig)
            tbl = pd.DataFrame({"Bulan": months, "F&B (Rp)": f_v, "Playground (Rp)": p_v,
                                "Total (Rp)": [a + b for a, b in zip(f_v, p_v)]})
            st.markdown("###### Ringkasan bulanan")
            st.dataframe(tbl, use_container_width=True, hide_index=True,
                column_config={
                    "F&B (Rp)":        st.column_config.NumberColumn(format="localized"),
                    "Playground (Rp)": st.column_config.NumberColumn(format="localized"),
                    "Total (Rp)":      st.column_config.NumberColumn(format="localized"),
                })

    st.markdown("###### Ekspor")
    ec = st.columns(4)
    s_, e_ = dr[0], dr[1]
    with ec[0]:
        st.download_button("🖨️ Print / PDF", generate_master_html(df_fnb, df_pg, s_, e_),
                           f"grove_master_{s_}_{e_}.html", "text/html", use_container_width=True)
    with ec[1]:
        st.download_button("📊 XLSX", generate_master_xlsx(df_fnb, df_pg, s_, e_),
                           f"grove_master_{s_}_{e_}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with ec[2]:
        raw = pd.concat([df_fnb.assign(sumber="F&B"), df_pg.assign(sumber="Playground")],
                        ignore_index=True) if not (fnb_empty or pg_empty) else (
              df_fnb.assign(sumber="F&B") if not fnb_empty else df_pg.assign(sumber="Playground"))
        st.download_button("⬇️ Data mentah (CSV)", raw.to_csv(index=False).encode(),
                           "grove_master_raw.csv", "text/csv", use_container_width=True)
    with ec[3]:
        summ = pd.DataFrame({"segmen": ["F&B", "Playground"], "nett_sales": [fnb_nett, pg_nett]})
        st.download_button("⬇️ Ringkasan (CSV)", summ.to_csv(index=False).encode(),
                           "grove_master_summary.csv", "text/csv", use_container_width=True)


def get_db():
    return SupabaseDB()


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon=APP_ICON, layout="wide", initial_sidebar_state="expanded")
    apply_custom_css()

    db = get_db()

    # First contact with the database. Failing here almost always means the
    # secrets are wrong or schema.sql was never run, so say that plainly
    # instead of surfacing a raw PostgREST error.
    try:
        has_users = db.has_users()
    except Exception as e:
        st.error(
            "**Tidak bisa terhubung ke database.**\n\n"
            "Periksa dua hal:\n\n"
            "1. `url` dan `service_key` di bagian `[supabase]` pada Streamlit Secrets "
            "sudah benar — pastikan memakai kunci **service_role**, bukan `anon`.\n"
            "2. `supabase/schema.sql` sudah dijalankan di Supabase Studio → SQL Editor."
        )
        st.caption(f"Detail teknis — {type(e).__name__}: {e}")
        return

    # First-time setup
    if not has_users:
        page_setup()
        return

    # Auth check
    user = check_auth()
    if not user:
        page_login()
        return

    st.session_state["user"] = user

    sel = render_sidebar()
    if "Master Dashboard" in sel: page_master_dashboard()
    elif "Dashboard F&B" in sel: page_dashboard_fnb()
    elif "Dashboard Playground" in sel: page_dashboard_playground()
    elif "Upload F&B" in sel: page_upload_esb()
    elif "Upload Playground" in sel: page_upload_playground()
    # Must precede the "Tenant" test below -- "Performa Tenant" contains it.
    elif "Performa" in sel: page_performance()
    elif "Unit" in sel: page_units()
    elif "Tenant" in sel: page_tenants()
    elif "User" in sel: page_users()
    elif "Log" in sel: page_upload_log()


if __name__ == "__main__":
    main()
