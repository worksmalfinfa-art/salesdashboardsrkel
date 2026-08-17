-- ============================================================================
-- GROVE Sales Analytics — Supabase (PostgreSQL) schema
-- ----------------------------------------------------------------------------
-- Replaces the 5 Google Sheets tabs: Users, Tenants, SalesData,
-- PlaygroundSales, UploadLog.
--
-- Run once in Supabase Studio → SQL Editor → New query → Run.
-- Safe to re-run: every statement is idempotent.
-- ============================================================================


-- ============================================================================
-- users  (was: "Users" sheet)
-- ============================================================================
create table if not exists public.users (
    email          text primary key,
    display_name   text        not null,
    role           text        not null default 'Viewer',
    tenant_access  text        not null default 'ALL',
    created_by     text,
    created_at     timestamptz not null default now(),
    is_active      boolean     not null default true,

    constraint users_role_valid
        check (role in ('Super Admin', 'Admin', 'Manager', 'Viewer'))
);


-- ============================================================================
-- tenants  (was: "Tenants" sheet)
-- ----------------------------------------------------------------------------
-- esb_branch_name must match the ESB export's "Branch Name" column exactly —
-- it is the join key used when importing ESB files.
-- ============================================================================
create table if not exists public.tenants (
    tenant_id        text primary key,
    tenant_name      text        not null unique,
    esb_branch_name  text        not null unique,
    created_at       timestamptz not null default now(),
    is_active        boolean     not null default true
);


-- ============================================================================
-- sales_data  (was: "SalesData" sheet)  — F&B hourly sales from ESB POS
-- ----------------------------------------------------------------------------
-- The FK to tenants(tenant_name) does two jobs that app.py previously did by
-- hand, row by row, over the Sheets API:
--   ON UPDATE CASCADE  — renaming a tenant rewrites its sales rows automatically
--   ON DELETE CASCADE  — deleting a tenant removes its sales rows automatically
--
-- The UNIQUE constraint enforces the (tenant, date, hour) de-duplication that
-- get_existing_keys() checked in Python, so a double upload can no longer
-- create duplicate rows even if the UI check is bypassed.
-- ============================================================================
create table if not exists public.sales_data (
    id              bigint generated always as identity primary key,
    tenant_name     text        not null
                    references public.tenants (tenant_name)
                    on update cascade on delete cascade,
    sales_date      date        not null,
    sales_hour      text        not null,
    pax_total       integer     not null default 0,
    subtotal        bigint      not null default 0,
    discount_total  bigint      not null default 0,
    nett_sales      bigint      not null default 0,
    uploaded_by     text,
    uploaded_at     timestamptz not null default now(),

    constraint sales_data_unique_slot
        unique (tenant_name, sales_date, sales_hour)
);

create index if not exists sales_data_date_idx        on public.sales_data (sales_date);
create index if not exists sales_data_tenant_date_idx on public.sales_data (tenant_name, sales_date);


-- ============================================================================
-- playground_sales  (was: "PlaygroundSales" sheet) — Twist N' Turns POS
-- ----------------------------------------------------------------------------
-- order_id is the primary key, so de-duplication is handled by the database.
-- Re-uploading an overlapping CSV is a no-op for rows already present.
-- ============================================================================
create table if not exists public.playground_sales (
    order_id         text primary key,
    amount           bigint      not null default 0,
    nett_sales       bigint      not null default 0,
    tax_amount       bigint      not null default 0,
    customer_name    text,
    sales_date       date        not null,
    child_total      integer     not null default 0,
    companion_total  integer     not null default 0,
    uploaded_by      text,
    uploaded_at      timestamptz not null default now()
);

create index if not exists playground_sales_date_idx on public.playground_sales (sales_date);


-- ============================================================================
-- upload_log  (was: "UploadLog" sheet) — audit trail
-- ============================================================================
create table if not exists public.upload_log (
    upload_id    text primary key,
    tenant_name  text,
    file_name    text,
    rows_added   integer     not null default 0,
    uploaded_by  text,
    uploaded_at  timestamptz not null default now()
);

create index if not exists upload_log_uploaded_at_idx on public.upload_log (uploaded_at desc);


-- ============================================================================
-- Row Level Security
-- ----------------------------------------------------------------------------
-- app.py connects with the service_role key, which bypasses RLS. RLS is enabled
-- here so that the tables are not world-readable if the anon key is ever
-- exposed, or if a public tenant-facing client is added later.
--
-- Access control for the app itself is enforced in check_auth() via the
-- users.role / users.tenant_access columns.
-- ============================================================================
alter table public.users            enable row level security;
alter table public.tenants          enable row level security;
alter table public.sales_data       enable row level security;
alter table public.playground_sales enable row level security;
alter table public.upload_log       enable row level security;

-- No permissive policies are defined: the anon and authenticated roles get no
-- access at all. Add policies here only when a non-service_role client exists.


-- ============================================================================
-- Seed: tenant registry
-- ----------------------------------------------------------------------------
-- esb_branch_name values below are PLACEHOLDERS. Replace each one with the
-- exact string in the "Branch Name" column (cell A15 onward) of that tenant's
-- ESB summary export before importing any ESB file — the import matches on
-- this string and silently finds nothing if it differs.
--
-- "Ruuang Kopi Cibis" is taken from a real ESB export header and is the only
-- value here with any evidence behind it; still verify it against the summary
-- report, which may format the name differently.
-- ============================================================================
insert into public.tenants (tenant_id, tenant_name, esb_branch_name) values
    ('T001', 'Omonyo',          'Omonyo'),
    ('T002', 'Mr Roastman',     'Mr Roastman'),
    ('T003', 'Sliced Pizzeria', 'Sliced Pizzeria'),
    ('T004', 'Ruuang Kopi',     'Ruuang Kopi Cibis'),
    ('T005', 'DOD Cafe',        'DOD Cafe')
on conflict (tenant_id) do nothing;
