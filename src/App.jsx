-- ============================================================
-- HoursLedger database schema — v7 (cost centers, departments, rental expenses, non-chargeable activities, employee departure date, billing follow-up)
-- Run this once, in full, in your Supabase project's SQL Editor.
-- Safe to re-run: every statement is idempotent.
-- ============================================================

create extension if not exists "pgcrypto";

-- ---------- profiles: one row per employee/manager ----------
create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  name text not null,
  email text not null,
  title text,
  role text not null default 'employee' check (role in ('employee','admin')),
  weekly_hours numeric not null default 40,
  annual_leave_days numeric not null default 25,
  active boolean not null default true,
  created_at timestamptz not null default now()
);
alter table public.profiles add column if not exists title text;
alter table public.profiles add column if not exists weekly_hours numeric not null default 40;
alter table public.profiles add column if not exists annual_leave_days numeric not null default 25;
-- Department the person belongs to — drives where their labor cost lands in
-- the P&L (Accounting Team = accounting/tax/payroll/other by logged hours;
-- Management Team = director/finance; Rental = rental).
alter table public.profiles add column if not exists department text not null default 'accounting_team';
alter table public.profiles drop constraint if exists profiles_department_check;
alter table public.profiles add constraint profiles_department_check
  check (department in ('accounting_team','management_team','rental'));
-- Departure date: like a client's end_date, this only stops future periods
-- from counting the person's cost — it never erases past periods. Active/
-- inactive still separately controls whether they can be assigned new work.
alter table public.profiles add column if not exists end_date date;
-- Superseded by department (only exists on installs that ran the interim v3).
alter table public.profiles drop constraint if exists profiles_cost_center_check;
alter table public.profiles drop column if exists cost_center;
alter table public.profiles drop column if exists cost_center_note;

-- ---------- employee labor cost history ----------
-- Each row is "these were the numbers from this date onward" — never
-- overwritten, so past periods stay accurate even after a raise.
create table if not exists public.employee_cost_history (
  id uuid primary key default gen_random_uuid(),
  employee_id uuid not null references public.profiles(id) on delete cascade,
  effective_date date not null,
  gross_salary numeric not null default 0,
  social_security numeric not null default 0,
  ticket_restaurant numeric not null default 0,
  insurance numeric not null default 0,
  other_cost numeric not null default 0,
  created_at timestamptz not null default now()
);
create index if not exists employee_cost_history_employee_idx on public.employee_cost_history(employee_id, effective_date);

-- ---------- clients ----------
create table if not exists public.clients (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  fixed_fee numeric,
  active boolean not null default true,
  created_at timestamptz not null default now()
);
-- Older installs may have these from before the fee-history redesign —
-- drop them if present; fixed_fee + the history/extra-fee tables replace them.
alter table public.clients drop column if exists monthly_hours;
alter table public.clients drop column if exists hourly_rate;

-- ---------- client fee & allocation history ----------
-- Same "never overwrite the past" pattern as employee cost history.
-- Allocation is split by service category for workload tracking.
create table if not exists public.client_fee_history (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  effective_date date not null,
  fixed_fee numeric,
  alloc_accounting numeric not null default 0,
  alloc_tax numeric not null default 0,
  alloc_payroll numeric not null default 0,
  alloc_other numeric not null default 0,
  created_at timestamptz not null default now()
);
create index if not exists client_fee_history_client_idx on public.client_fee_history(client_id, effective_date);
-- Cost-center split of the fixed fee, in euro per month. The app requires the
-- seven amounts to add up to fixed_fee; rows with no split (all zero) are
-- reported as "Unallocated" on the dashboard until you edit them.
alter table public.client_fee_history add column if not exists cc_accounting numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_tax numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_payroll numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_rental numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_director numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_finance numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_other numeric not null default 0;
alter table public.client_fee_history add column if not exists cc_other_note text;

-- ---------- client extra fees: one-off charges for a specific month ----------
create table if not exists public.client_extra_fees (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  month char(7) not null, -- 'YYYY-MM'
  amount numeric not null check (amount > 0),
  note text,
  created_at timestamptz not null default now()
);
create index if not exists client_extra_fees_client_idx on public.client_extra_fees(client_id, month);
-- Each extra fee is tagged with one cost center (note required for 'other').
alter table public.client_extra_fees add column if not exists cost_center text not null default 'other';
alter table public.client_extra_fees drop constraint if exists client_extra_fees_cost_center_check;
alter table public.client_extra_fees add constraint client_extra_fees_cost_center_check
  check (cost_center in ('accounting','tax','payroll','other','director','finance','rental'));

-- ---------- rental properties & their expenses ----------
-- Non-labor cost of the Rental department: rent, utilities and common
-- expenses, booked per property and per month ('YYYY-MM').
create table if not exists public.rental_properties (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  address text,
  note text,
  active boolean not null default true,
  created_at timestamptz not null default now()
);
-- kind = 'recurring': amount per month from from_month to to_month (null =
-- open-ended). kind = 'extra': one-off amount booked to a single month.
create table if not exists public.rental_expenses (
  id uuid primary key default gen_random_uuid(),
  property_id uuid not null references public.rental_properties(id) on delete restrict,
  kind text not null default 'extra' check (kind in ('recurring','extra')),
  month char(7),      -- 'YYYY-MM', extra only
  from_month char(7), -- 'YYYY-MM', recurring only
  to_month char(7),   -- 'YYYY-MM' or null = open-ended, recurring only
  category text not null check (category in ('rent','utilities','common_expenses')),
  amount numeric not null, -- recurring: > 0 per month; extra: any non-zero (negative = discount/rebate)
  note text,
  created_at timestamptz not null default now()
);
-- Upgrade installs that created the first (month-only) version of the table.
alter table public.rental_expenses add column if not exists kind text not null default 'extra';
alter table public.rental_expenses add column if not exists from_month char(7);
alter table public.rental_expenses add column if not exists to_month char(7);
alter table public.rental_expenses alter column month drop not null;
alter table public.rental_expenses drop constraint if exists rental_expenses_kind_check;
alter table public.rental_expenses add constraint rental_expenses_kind_check check (kind in ('recurring','extra'));
alter table public.rental_expenses drop constraint if exists rental_expenses_amount_check;
alter table public.rental_expenses add constraint rental_expenses_amount_check check (
  (kind = 'recurring' and amount > 0) or (kind = 'extra' and amount <> 0)
);
alter table public.rental_expenses drop constraint if exists rental_expenses_period_check;
alter table public.rental_expenses add constraint rental_expenses_period_check check (
  (kind = 'extra' and month is not null) or
  (kind = 'recurring' and from_month is not null and (to_month is null or to_month >= from_month))
);
create index if not exists rental_expenses_property_idx on public.rental_expenses(property_id, kind, month, from_month);

-- ---------- time entries ----------
create table if not exists public.time_entries (
  id uuid primary key default gen_random_uuid(),
  employee_id uuid not null references public.profiles(id) on delete cascade,
  client_id uuid references public.clients(id) on delete restrict, -- null for internal activities
  entry_date date not null,
  hours numeric not null check (hours > 0),
  category text not null default 'other' check (category in ('accounting','tax','payroll','other')),
  note text,
  created_at timestamptz not null default now()
);
alter table public.time_entries add column if not exists category text not null default 'other';
alter table public.time_entries add column if not exists note text;
-- Internal, non-chargeable activities: an entry is EITHER for a client
-- (client_id set, activity null) OR for an activity (activity set, client_id
-- null). Activity entries carry category 'other' and no revenue.
alter table public.time_entries alter column client_id drop not null;
alter table public.time_entries add column if not exists activity text;
alter table public.time_entries drop constraint if exists time_entries_activity_check;
alter table public.time_entries add constraint time_entries_activity_check
  check (activity is null or activity in ('business_development','training','client_communication','other_non_chargeable'));
alter table public.time_entries drop constraint if exists time_entries_target_check;
alter table public.time_entries add constraint time_entries_target_check
  check ((client_id is not null and activity is null) or (client_id is null and activity is not null));
create index if not exists time_entries_activity_idx on public.time_entries(activity) where activity is not null;
create index if not exists time_entries_employee_date_idx on public.time_entries(employee_id, entry_date);
create index if not exists time_entries_client_date_idx on public.time_entries(client_id, entry_date);

-- ---------- locked (submitted) weeks ----------
create table if not exists public.locked_weeks (
  employee_id uuid not null references public.profiles(id) on delete cascade,
  week_start date not null,
  locked_at timestamptz not null default now(),
  primary key (employee_id, week_start)
);

-- ---------- billing follow-up (παρακολούθηση τιμολόγησης & εξόφλησης) ----------
-- Invoices are issued outside the app. The app derives every billable line
-- itself (one per client per month of fixed fee, one per extra fee) and
-- this table only stores the follow-up status of a line: when it was
-- invoiced and when it was paid. No row = not invoiced yet.
create table if not exists public.billing_status (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  kind text not null check (kind in ('fixed','extra')),
  month char(7) not null, -- 'YYYY-MM'
  extra_fee_id uuid references public.client_extra_fees(id) on delete cascade, -- extra only
  invoiced_date date,
  invoice_ref text,
  paid_date date,
  paid_amount numeric, -- null = paid in full
  note text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check ((kind = 'fixed' and extra_fee_id is null) or (kind = 'extra' and extra_fee_id is not null))
);
create unique index if not exists billing_status_fixed_key on public.billing_status(client_id, month) where kind = 'fixed';
create unique index if not exists billing_status_extra_key on public.billing_status(extra_fee_id) where kind = 'extra';

-- The earlier draft of this feature created invoice tables; they are
-- replaced by billing_status. Harmless if they were never created.
drop table if exists public.invoice_payments;
drop table if exists public.invoices;

-- ---------- app settings: small shared key/value store ----------
-- Used for things like the admin's Dashboard "Customize" layout choices —
-- not sensitive, just UI preferences shared across the team.
create table if not exists public.app_settings (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

-- ---------- leave requests ----------
create table if not exists public.leave_requests (
  id uuid primary key default gen_random_uuid(),
  employee_id uuid not null references public.profiles(id) on delete cascade,
  start_date date not null,
  end_date date not null,
  type text not null check (type in ('annual','sick','medical','other_legal','extra_acco')),
  note text,
  status text not null default 'pending' check (status in ('pending','approved','rejected')),
  decided_by uuid references public.profiles(id),
  decided_at date,
  created_at timestamptz not null default now()
);
create index if not exists leave_requests_employee_idx on public.leave_requests(employee_id, start_date);

-- ============================================================
-- Auto-create a profile row whenever someone signs up.
-- ============================================================
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
  insert into public.profiles (id, name, email, role, active)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'name', split_part(new.email, '@', 1)),
    new.email,
    'employee',
    true
  );
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- ============================================================
-- Helper: is the current logged-in user an active admin?
-- ============================================================
create or replace function public.is_admin()
returns boolean
language sql
security definer set search_path = public
stable
as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'admin' and active = true
  );
$$;

-- ============================================================
-- Row Level Security
-- ============================================================
alter table public.profiles enable row level security;
alter table public.employee_cost_history enable row level security;
alter table public.clients enable row level security;
alter table public.client_fee_history enable row level security;
alter table public.client_extra_fees enable row level security;
alter table public.time_entries enable row level security;
alter table public.locked_weeks enable row level security;
alter table public.leave_requests enable row level security;
alter table public.app_settings enable row level security;
alter table public.billing_status enable row level security;
alter table public.rental_properties enable row level security;
alter table public.rental_expenses enable row level security;

-- profiles
drop policy if exists "profiles_select_all" on public.profiles;
create policy "profiles_select_all" on public.profiles
  for select using (auth.uid() is not null);

drop policy if exists "profiles_update_self" on public.profiles;
create policy "profiles_update_self" on public.profiles
  for update using (auth.uid() = id) with check (auth.uid() = id);

drop policy if exists "profiles_update_admin" on public.profiles;
create policy "profiles_update_admin" on public.profiles
  for update using (public.is_admin());

-- employee cost history: salary data — admins only, in every direction.
drop policy if exists "cost_history_admin_all" on public.employee_cost_history;
create policy "cost_history_admin_all" on public.employee_cost_history
  for all using (public.is_admin()) with check (public.is_admin());

-- clients
drop policy if exists "clients_select_all" on public.clients;
create policy "clients_select_all" on public.clients
  for select using (auth.uid() is not null);

drop policy if exists "clients_insert_admin" on public.clients;
create policy "clients_insert_admin" on public.clients
  for insert with check (public.is_admin());

drop policy if exists "clients_update_admin" on public.clients;
create policy "clients_update_admin" on public.clients
  for update using (public.is_admin());

-- client fee history & extra fees: billing terms — admins only.
drop policy if exists "fee_history_admin_all" on public.client_fee_history;
create policy "fee_history_admin_all" on public.client_fee_history
  for all using (public.is_admin()) with check (public.is_admin());

drop policy if exists "extra_fees_admin_all" on public.client_extra_fees;
create policy "extra_fees_admin_all" on public.client_extra_fees
  for all using (public.is_admin()) with check (public.is_admin());

-- time entries: employees manage their own; admins manage everyone's.
drop policy if exists "time_entries_select" on public.time_entries;
create policy "time_entries_select" on public.time_entries
  for select using (employee_id = auth.uid() or public.is_admin());

drop policy if exists "time_entries_insert" on public.time_entries;
create policy "time_entries_insert" on public.time_entries
  for insert with check (employee_id = auth.uid() or public.is_admin());

drop policy if exists "time_entries_update" on public.time_entries;
create policy "time_entries_update" on public.time_entries
  for update using (employee_id = auth.uid() or public.is_admin());

drop policy if exists "time_entries_delete" on public.time_entries;
create policy "time_entries_delete" on public.time_entries
  for delete using (employee_id = auth.uid() or public.is_admin());

-- locked weeks
drop policy if exists "locked_weeks_select" on public.locked_weeks;
create policy "locked_weeks_select" on public.locked_weeks
  for select using (employee_id = auth.uid() or public.is_admin());

drop policy if exists "locked_weeks_insert" on public.locked_weeks;
create policy "locked_weeks_insert" on public.locked_weeks
  for insert with check (employee_id = auth.uid() or public.is_admin());

drop policy if exists "locked_weeks_delete" on public.locked_weeks;
create policy "locked_weeks_delete" on public.locked_weeks
  for delete using (public.is_admin());
-- Note: only admins can delete (reopen) a locked week. Employees can still
-- lock their own week (insert, policy above) but not unlock it themselves —
-- enforced here at the database level, not just hidden in the UI.

-- leave requests: employees manage their own; admins see/decide on all.
drop policy if exists "leave_requests_select" on public.leave_requests;
create policy "leave_requests_select" on public.leave_requests
  for select using (employee_id = auth.uid() or public.is_admin());

drop policy if exists "leave_requests_insert" on public.leave_requests;
create policy "leave_requests_insert" on public.leave_requests
  for insert with check (employee_id = auth.uid() or public.is_admin());

drop policy if exists "leave_requests_update" on public.leave_requests;
create policy "leave_requests_update" on public.leave_requests
  for update using (employee_id = auth.uid() or public.is_admin());

drop policy if exists "leave_requests_delete" on public.leave_requests;
create policy "leave_requests_delete" on public.leave_requests
  for delete using (employee_id = auth.uid() or public.is_admin());

-- billing follow-up: admins only.
drop policy if exists "billing_status_admin_all" on public.billing_status;
create policy "billing_status_admin_all" on public.billing_status
  for all using (public.is_admin()) with check (public.is_admin());

-- rental properties & expenses: cost data — admins only.
drop policy if exists "rental_properties_admin_all" on public.rental_properties;
create policy "rental_properties_admin_all" on public.rental_properties
  for all using (public.is_admin()) with check (public.is_admin());

drop policy if exists "rental_expenses_admin_all" on public.rental_expenses;
create policy "rental_expenses_admin_all" on public.rental_expenses
  for all using (public.is_admin()) with check (public.is_admin());

-- app settings: readable by any signed-in user, writable by admins only.
drop policy if exists "app_settings_select" on public.app_settings;
create policy "app_settings_select" on public.app_settings
  for select using (auth.uid() is not null);

drop policy if exists "app_settings_upsert_admin" on public.app_settings;
create policy "app_settings_upsert_admin" on public.app_settings
  for insert with check (public.is_admin());

drop policy if exists "app_settings_update_admin" on public.app_settings;
create policy "app_settings_update_admin" on public.app_settings
  for update using (public.is_admin());
