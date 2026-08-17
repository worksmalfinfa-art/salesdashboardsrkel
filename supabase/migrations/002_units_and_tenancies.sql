-- ============================================================================
-- GROVE Sales Analytics — migration 002
-- Separates the physical space from the brand that occupies it
-- ----------------------------------------------------------------------------
-- Until now a tenant WAS its name, and sales rows pointed at that name. That
-- made two ordinary events impossible to record correctly:
--
--   Renaming a brand rewrote its whole sales history (ON UPDATE CASCADE), and
--   replacing the occupant of a unit had no representation at all — you either
--   corrupted the outgoing brand's history by renaming it, or created a
--   disconnected new tenant and lost the fact that it is the same space.
--
-- This migration introduces three entities where there was one:
--
--   units      the physical space   — permanent
--   tenants    the brand            — comes and goes
--   tenancies  who occupied what, between which dates
--
-- Sales stay attached to the BRAND (via a stable tenant_id, not the name), so
-- brand history survives both renames and hand-overs. The unit a sale belongs
-- to is derived by looking up which tenancy covered that date.
--
-- Run once in Supabase Studio → SQL Editor. Safe to re-run.
-- ============================================================================


-- ============================================================================
-- 1. units — the physical space
-- ============================================================================
create table if not exists public.units (
    unit_id     text primary key,
    unit_code   text        not null unique,   -- e.g. 'Slot DOD', 'A-3'
    unit_name   text,                          -- optional description
    floor       text,
    area_sqm    numeric(10,2),                 -- enables sales per m² later
    is_active   boolean     not null default true,
    created_at  timestamptz not null default now()
);


-- ============================================================================
-- 2. tenants — the brand
-- ----------------------------------------------------------------------------
-- New columns make the F&B / Playground split a property of the data instead
-- of a branch in the code, so adding a tenant of either kind needs no edit.
-- esb_branch_name becomes nullable: a Playground tenant has no ESB branch.
-- ============================================================================
alter table public.tenants
    add column if not exists category  text not null default 'F&B',
    add column if not exists pos_type  text not null default 'esb';

do $$ begin
    alter table public.tenants
        add constraint tenants_pos_type_valid check (pos_type in ('esb', 'playground'));
exception when duplicate_object then null;
end $$;

alter table public.tenants alter column esb_branch_name drop not null;


-- ============================================================================
-- 3. tenancies — who occupied which unit, when
-- ----------------------------------------------------------------------------
-- end_date NULL means "still occupying". Replacing a tenant is therefore:
-- set end_date on the current row, insert a new one. Nothing else moves.
-- ============================================================================
create table if not exists public.tenancies (
    tenancy_id  bigint generated always as identity primary key,
    unit_id     text not null references public.units (unit_id)   on delete cascade,
    tenant_id   text not null references public.tenants (tenant_id) on delete cascade,
    start_date  date not null,
    end_date    date,
    notes       text,
    created_at  timestamptz not null default now(),

    constraint tenancies_dates_ordered
        check (end_date is null or end_date >= start_date)
);

create index if not exists tenancies_unit_idx   on public.tenancies (unit_id);
create index if not exists tenancies_tenant_idx on public.tenancies (tenant_id);

-- Two tenants cannot occupy the same unit on the same day. Enforced by the
-- database rather than the UI, because an overlap here silently double-counts
-- that unit's revenue in every per-unit report.
do $$ begin
    create extension if not exists btree_gist;
    begin
        alter table public.tenancies add constraint tenancies_no_overlap
            exclude using gist (
                unit_id with =,
                daterange(start_date, coalesce(end_date, 'infinity'::date), '[]') with &&
            );
    exception when duplicate_object then null;
    end;
exception when others then
    raise notice 'btree_gist unavailable — overlap protection skipped (%).', sqlerrm;
end $$;


-- ============================================================================
-- 4. sales_data — repoint from tenant_name to tenant_id
-- ----------------------------------------------------------------------------
-- Done in place so existing rows are preserved. The backfill matches on the
-- name that is being retired, which is why it must run before the old column
-- is dropped.
-- ============================================================================
alter table public.sales_data add column if not exists tenant_id text;

update public.sales_data s
   set tenant_id = t.tenant_id
  from public.tenants t
 where t.tenant_name = s.tenant_name
   and s.tenant_id is null;

-- Any row whose tenant_name matches no tenant would block the NOT NULL below.
-- Surface it loudly rather than failing on a constraint the reader can't place.
do $$
declare orphans int;
begin
    select count(*) into orphans from public.sales_data where tenant_id is null;
    if orphans > 0 then
        raise exception 'ABORT: % sales rows have a tenant_name matching no tenant. '
                        'Fix those rows or add the missing tenants, then re-run.', orphans;
    end if;
end $$;

do $$ begin
    alter table public.sales_data drop constraint if exists sales_data_tenant_name_fkey;
    alter table public.sales_data drop constraint if exists sales_data_unique_slot;
exception when undefined_object then null;
end $$;

alter table public.sales_data alter column tenant_id set not null;

do $$ begin
    alter table public.sales_data
        add constraint sales_data_tenant_fkey
        foreign key (tenant_id) references public.tenants (tenant_id) on delete cascade;
exception when duplicate_object then null;
end $$;

do $$ begin
    alter table public.sales_data
        add constraint sales_data_unique_slot unique (tenant_id, sales_date, sales_hour);
exception when duplicate_object then null;
end $$;

alter table public.sales_data drop column if exists tenant_name;

create index if not exists sales_data_tenant_date_idx on public.sales_data (tenant_id, sales_date);


-- ============================================================================
-- 5. playground_sales — attach to a tenant like everything else
-- ----------------------------------------------------------------------------
-- Playground was a hard-coded second world. Giving it a tenant_id lets the
-- master dashboard treat it as one more tenant with a different POS type.
-- ============================================================================
alter table public.playground_sales add column if not exists tenant_id text;

do $$ begin
    alter table public.playground_sales
        add constraint playground_tenant_fkey
        foreign key (tenant_id) references public.tenants (tenant_id) on delete set null;
exception when duplicate_object then null;
end $$;


-- ============================================================================
-- 6. targets — monthly revenue target per tenant, editable in the app
-- ============================================================================
create table if not exists public.targets (
    target_id     bigint generated always as identity primary key,
    tenant_id     text not null references public.tenants (tenant_id) on delete cascade,
    period_month  date not null,            -- always the 1st of the month
    target_nett   bigint not null default 0,
    updated_by    text,
    updated_at    timestamptz not null default now(),

    constraint targets_unique_period unique (tenant_id, period_month)
);


-- ============================================================================
-- 7. Seed: a Playground tenant, then a unit and an open tenancy per tenant
-- ----------------------------------------------------------------------------
-- Every existing tenant gets a unit named after it and an open-ended tenancy,
-- so nothing is orphaned on day one. Rename the units to your real slot codes
-- in the app afterwards — the unit is the permanent thing, so 'Slot DOD' is a
-- better code than the name of whoever happens to occupy it today.
-- ============================================================================
insert into public.tenants (tenant_id, tenant_name, esb_branch_name, category, pos_type)
values ('T900', 'Twist N'' Turns Playground', null, 'Playground', 'playground')
on conflict (tenant_id) do nothing;

update public.playground_sales set tenant_id = 'T900' where tenant_id is null;

insert into public.units (unit_id, unit_code)
select 'U' || substring(t.tenant_id from 2), t.tenant_name
  from public.tenants t
 where not exists (
       select 1 from public.units u where u.unit_id = 'U' || substring(t.tenant_id from 2))
on conflict (unit_code) do nothing;

insert into public.tenancies (unit_id, tenant_id, start_date, notes)
select 'U' || substring(t.tenant_id from 2), t.tenant_id, date '2025-01-01',
       'Auto-created by migration 002'
  from public.tenants t
 where exists (select 1 from public.units u where u.unit_id = 'U' || substring(t.tenant_id from 2))
   and not exists (select 1 from public.tenancies tn where tn.tenant_id = t.tenant_id)
on conflict do nothing;


-- ============================================================================
-- 8. v_sales_enriched — sales with the brand and the unit it belonged to
-- ----------------------------------------------------------------------------
-- The tenancy join is by DATE, so a sale made before a hand-over is credited
-- to the unit's occupant at that time, not to whoever occupies it now. This is
-- the whole point of the model, and doing it here means the app never has to
-- get the date logic right twice.
-- ============================================================================
create or replace view public.v_sales_enriched as
select s.id,
       s.tenant_id,
       t.tenant_name,
       t.category,
       s.sales_date,
       s.sales_hour,
       s.pax_total,
       s.subtotal,
       s.discount_total,
       s.nett_sales,
       u.unit_id,
       u.unit_code
  from public.sales_data s
  join public.tenants  t on t.tenant_id = s.tenant_id
  left join public.tenancies tn
         on tn.tenant_id = s.tenant_id
        and s.sales_date >= tn.start_date
        and s.sales_date <= coalesce(tn.end_date, 'infinity'::date)
  left join public.units u on u.unit_id = tn.unit_id;


-- ============================================================================
-- 9. Row Level Security on the new tables
-- ============================================================================
alter table public.units     enable row level security;
alter table public.tenancies enable row level security;
alter table public.targets   enable row level security;
