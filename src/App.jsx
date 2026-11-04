-- ============================================================
-- HoursLedger database schema — v11 (agreement: other charges text, operating expenses, annual & start-up fees on the client card, cost centers, departments, rental expenses, non-chargeable activities, employee departure date, billing follow-up, CRM phase 1: client card, contacts, notes, tasks, lifecycle, staged onboarding checklist, AML due diligence, client documents & agreement, AML compliance toolkit)
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

-- ---------- CRM: client profile, contacts, notes, tasks ----------
-- Client profile fields (ΑΦΜ, ΔΟΥ, legal form, address, contact details,
-- account manager at the firm). Readable by any signed-in user (the UI
-- only shows the full card to managers); writable by managers.
alter table public.clients add column if not exists afm text;
alter table public.clients add column if not exists doy text;
alter table public.clients add column if not exists legal_form text;
alter table public.clients add column if not exists address text;
alter table public.clients add column if not exists email text;
alter table public.clients add column if not exists phone text;
alter table public.clients add column if not exists account_manager_id uuid references public.profiles(id) on delete set null;
alter table public.clients add column if not exists general_notes text;
-- Legal representative (also used by due diligence and the agreement).
alter table public.clients add column if not exists representative_name text;
alter table public.clients add column if not exists representative_afm text;
alter table public.clients add column if not exists representative_id_type text;
alter table public.clients add column if not exists representative_id_number text;
alter table public.clients add column if not exists representative_email text;
alter table public.clients add column if not exists representative_phone text;

-- Client lifecycle: lead (acquisition) → onboarding (setup checklist in
-- progress) → live (all onboarding tasks done) → ended. Existing clients
-- default to 'live'. go_live_date is set when onboarding completes.
alter table public.clients add column if not exists lifecycle text not null default 'live';
alter table public.clients drop constraint if exists clients_lifecycle_check;
alter table public.clients add constraint clients_lifecycle_check check (lifecycle in ('lead','onboarding','live','ended'));
alter table public.clients add column if not exists go_live_date date;
alter table public.clients add column if not exists lead_source text;
alter table public.clients add column if not exists expected_fee numeric;

-- Onboarding checklist template: the tasks created for every new client
-- when onboarding starts. Managed by admins in the app (Tasks page).
-- Each item belongs to one of three stages of the engagement:
--   setup    — operational setup of the client in the firm
--   contract — offer, engagement letter, GDPR, authorisations
--   aml      — customer due diligence under ν. 4557/2018
create table if not exists public.onboarding_templates (
  id uuid primary key default gen_random_uuid(),
  stage text not null default 'setup',
  sort_order int not null default 0,
  title text not null,
  description text,
  due_days int not null default 7, -- due = onboarding start + due_days
  active boolean not null default true,
  created_at timestamptz not null default now()
);
alter table public.onboarding_templates add column if not exists stage text not null default 'setup';
alter table public.onboarding_templates drop constraint if exists onboarding_templates_stage_check;
alter table public.onboarding_templates add constraint onboarding_templates_stage_check check (stage in ('setup','contract','aml','formation'));
-- applies_to: 'all' | 'capital' (Α.Ε., Ε.Π.Ε., Ι.Κ.Ε. only) | 'if_applicable' (created with that note)
alter table public.onboarding_templates add column if not exists applies_to text not null default 'all';
alter table public.onboarding_templates drop constraint if exists onboarding_templates_applies_check;
alter table public.onboarding_templates add constraint onboarding_templates_applies_check check (applies_to in ('all','capital','if_applicable'));
-- Default checklist. Items are inserted only if a template with the same
-- title does not already exist, so your own edits are kept.
insert into public.onboarding_templates (stage, sort_order, title, description, due_days, applies_to)
select v.* from (values
  ('formation', 1, 'Business set up στην Υπηρεσία μιας Στάσης (ΥΜΣ)', 'Σύσταση / έναρξη', 5, 'all'),
  ('formation', 2, 'Σύνταξη μισθωτηρίου συμβολαίου έδρας & υποβολή στην ΑΑΔΕ', '', 5, 'all'),
  ('formation', 3, 'Άνοιγμα τραπεζικού λογαριασμού', '', 7, 'all'),
  ('formation', 4, 'Δήλωση τραπεζικού λογαριασμού', 'ΑΑΔΕ / ΓΕΜΗ όπου απαιτείται', 7, 'all'),
  ('formation', 5, 'Εγγραφή ΕΦΚΑ μη μισθωτών', '', 10, 'all'),
  ('formation', 6, 'Απαλλαγή εισφορών για αλλοδαπό', 'Μόνο όπου εφαρμόζεται', 10, 'if_applicable'),
  ('formation', 7, 'Πιστοποίηση καταβολής εταιρικού κεφαλαίου', 'ΙΚΕ, ΑΕ, ΕΠΕ', 14, 'capital'),
  ('formation', 8, 'Set up ηλεκτρονικής τιμολόγησης, δηλώσεις παρόχων, δηλώσεις POS', '', 10, 'all'),
  ('formation', 9, 'Ενημέρωση για θέματα που απαιτούν μηχανικό στο myBusiness', '', 7, 'all'),
  ('formation', 10, 'Δήλωση εταιρικού site στο ΓΕΜΗ', 'ΙΚΕ, ΑΕ, ΕΠΕ', 14, 'capital'),
  ('aml', 1, 'Έγγραφα ΠΟΛ.1200/2018 & έντυπο KYC (πριν την ανάληψη)', 'Πίνακας Α/Β/Γ/Δ κατά περίπτωση· έντυπο KYC με ενημέρωση προσωπικών δεδομένων (άρ. 31 παρ. 2)· αν ενεργεί μέσω εξουσιοδοτημένου, ταυτοποίηση και εκείνου', 5, 'all'),
  ('aml', 2, 'Αποδεικτικό εγγραφής στο Κεντρικό Μητρώο Πραγματικών Δικαιούχων', 'Για νομικά πρόσωπα/εμπιστεύματα (άρ. 14 παρ. 1)· υπενθύμιση στον πελάτη αν δεν έχει καταχωρίσει', 5, 'all'),
  ('aml', 3, 'Εξακρίβωση πραγματικού δικαιούχου & χαρτογράφηση δομής', 'Μετοχολόγιο/κατάσταση εταίρων· >25% σε άλλη εταιρεία → επόμενο επίπεδο· ταυτοποίηση με Πίνακα Α· αν δεν προκύψει φυσικό πρόσωπο, ανώτατο στέλεχος + αρχείο ενεργειών', 5, 'all'),
  ('aml', 4, 'Έλεγχος ΠΕΠ & κυρώσεων', 'Πελάτης, νόμιμος εκπρόσωπος, πραγματικοί δικαιούχοι, στενοί συγγενείς & συνεργάτες — ημερομηνία και αποτέλεσμα', 5, 'all'),
  ('aml', 5, 'Σημείωμα σκοπού & φύσης της σχέσης (με πηγές)', 'Δραστηριότητα, υπηρεσίες μας, γιατί απευθύνθηκε, αναμενόμενη συναλλακτική εικόνα, πηγές (δήλωση, ΓΕΜΗ, έναρξη, ιστοσελίδα)', 5, 'all'),
  ('aml', 6, 'Βαθμολόγηση & κατάταξη κινδύνου', 'Άξονες Α–Δ στην καρτέλα → ΔΕΠ απλουστευμένη/συνήθης/αυξημένη· αν υψηλός, έγκριση ανώτερου στελέχους πριν την ανάληψη', 7, 'all'),
  ('aml', 7, 'Απόφαση ανάληψης ή αιτιολογημένη άρνηση', 'Αν δεν ολοκληρώνεται η ταυτοποίηση/σκοπός: άρνηση, τεκμηρίωση και εξέταση αναφοράς (άρ. 13 παρ. 2)', 7, 'all'),
  ('aml', 8, 'Άνοιγμα φακέλου & ημερομηνία επόμενης επικαιροποίησης', 'Χαμηλός: ανά διετία· μέτριος/υψηλός: ετησίως και σε κάθε μεταβολή· τήρηση 5 έτη μετά τη λήξη', 10, 'all'),
  ('contract', 1, 'Προσφορά υπηρεσιών & αποδοχή', 'Fee, περιεχόμενο υπηρεσιών, όροι πληρωμής', 3, 'all'),
  ('contract', 2, 'Σύμβαση παροχής λογιστικών υπηρεσιών / engagement letter', 'Υπογραφή και αρχειοθέτηση', 7, 'all'),
  ('contract', 3, 'Ενημέρωση & συγκατάθεση GDPR', 'Ενημέρωση επεξεργασίας προσωπικών δεδομένων', 7, 'all'),
  ('contract', 4, 'Εξουσιοδοτήσεις ΑΑΔΕ / ΕΦΚΑ / ΕΡΓΑΝΗ', 'myAADE εξουσιοδότηση λογιστή, κωδικοί όπου απαιτείται', 7, 'all'),
  ('setup', 1, 'Συμπλήρωση καρτέλας πελάτη', 'ΑΦΜ, ΔΟΥ, νομική μορφή, διεύθυνση, πρόσωπα επαφής, account manager', 3, 'all'),
  ('setup', 2, 'Ορισμός fee, allocation & cost centers στο GLORIA', 'Fees από τη λίστα πελατών', 3, 'all'),
  ('setup', 3, 'Πρόσβαση myDATA & setup στο λογιστικό πρόγραμμα', 'Εταιρεία, λογαριασμοί, σειρές παραστατικών', 10, 'all'),
  ('setup', 4, 'Παραλαβή αρχείου από προηγούμενο λογιστή', 'Ισοζύγια, μητρώο παγίων, εκκρεμότητες', 14, 'all'),
  ('setup', 5, 'Καταγραφή υποχρεώσεων & προθεσμιών', 'ΦΠΑ, ΑΠΔ, ΕΦΚΑ, δηλώσεις — ποιες ισχύουν για τον πελάτη', 10, 'all'),
  ('setup', 6, 'Άνοιγμα πελάτη στο Hyper (λογιστικό πρόγραμμα)', 'Δημιουργία εταιρείας/χρήσης, λογαριασμοί, σειρές παραστατικών', 10, 'all'),
  ('setup', 7, 'Άνοιγμα πελάτη στο tax system', 'Δημιουργία καρτέλας πελάτη με ΑΦΜ, ΔΟΥ και υποχρεώσεις', 10, 'all'),
  ('setup', 8, 'Set up ηλεκτρονικής τιμολόγησης', 'Πάροχος/πρόγραμμα έκδοσης e-τιμολογίων, σειρές παραστατικών, σύνδεση με myDATA', 10, 'all'),
  ('setup', 9, 'Δηλώσεις παρόχων', 'Δήλωση παρόχου ηλεκτρονικής τιμολόγησης και λοιπών παρόχων στο myAADE/myDATA', 10, 'all'),
  ('setup', 10, 'Δηλώσεις POS', 'Δήλωση POS / ΦΗΜ-ταμειακών και συνδεδεμένων τερματικών (όπου υπάρχουν)', 10, 'all')
) as v(stage, sort_order, title, description, due_days, applies_to)
where not exists (select 1 from public.onboarding_templates t where t.title = v.title);
-- Retire the first-version items that are now covered by the staged list.
update public.onboarding_templates set active = false
where title in ('Στοιχεία & έγγραφα KYC/AML','Σύμβαση συνεργασίας / engagement letter','Εξουσιοδότηση ΑΑΔΕ (myAADE) & κωδικοί','Πρόσβαση myDATA & λογιστικό πρόγραμμα','Ορισμός fee, allocation & account manager στο GLORIA')
  and stage = 'setup' and created_at < now() - interval '1 minute';

-- ---------- customer due diligence (KYC/AML, ν. 4557/2018) ----------
-- One record per client. Risk factors, documents and services are kept as
-- JSON so the checklist can evolve without schema changes. Admins only.
create table if not exists public.client_aml (
  client_id uuid primary key references public.clients(id) on delete cascade,
  client_type text not null default 'legal' check (client_type in ('legal','natural')),
  gemi_number text,
  incorporation_date date,
  representative_name text,
  representative_id_type text,
  representative_id_number text,
  representative_afm text,
  activity text,
  countries text,
  relationship_purpose text,
  services jsonb not null default '[]'::jsonb,
  source_of_funds text,
  pep_status text not null default 'none' check (pep_status in ('none','client','ubo','representative')),
  pep_details text,
  sanctions_checked_on date,
  sanctions_result text,
  risk_factors jsonb not null default '[]'::jsonb,
  risk_score int not null default 0,
  risk_level text not null default 'low' check (risk_level in ('low','medium','high')),
  risk_override text check (risk_override in ('low','medium','high')),
  risk_override_reason text,
  dd_level text not null default 'standard' check (dd_level in ('simplified','standard','enhanced')),
  documents jsonb not null default '{}'::jsonb,
  status text not null default 'not_started' check (status in ('not_started','in_progress','completed')),
  completed_on date,
  completed_by uuid references public.profiles(id) on delete set null,
  approved_by uuid references public.profiles(id) on delete set null,
  approved_on date,
  next_review_on date,
  notes text,
  updated_at timestamptz not null default now()
);
-- Fields added for the ΠΟΛ.1200/2018 / Taxheaven-guide structure.
alter table public.client_aml add column if not exists entity_type text; -- natural | capital | partnership | other
alter table public.client_aml add column if not exists info_sources text; -- where the KYC information came from
alter table public.client_aml add column if not exists proxy_name text;
alter table public.client_aml add column if not exists proxy_id_number text;
alter table public.client_aml add column if not exists decision text; -- accept | decline
alter table public.client_aml add column if not exists decision_reason text;
alter table public.client_aml add column if not exists ubo_actions_record text; -- actions taken when UBO could not be identified
alter table public.client_aml drop constraint if exists client_aml_status_check;
alter table public.client_aml add constraint client_aml_status_check check (status in ('not_started','in_progress','completed','declined'));
-- Risk sheet metadata (Φύλλο αξιολόγησης κινδύνου): per-factor documentation,
-- assessment type/date/version, Model B score, UBO-change tracking (άρ. 20).
alter table public.client_aml add column if not exists risk_factor_notes jsonb not null default '{}'::jsonb;
alter table public.client_aml add column if not exists assessment_type text not null default 'initial';
alter table public.client_aml drop constraint if exists client_aml_assessment_type_check;
alter table public.client_aml add constraint client_aml_assessment_type_check check (assessment_type in ('initial','annual','extraordinary'));
alter table public.client_aml add column if not exists assessment_date date;
alter table public.client_aml add column if not exists sheet_version int not null default 1;
alter table public.client_aml add column if not exists risk_score_b int not null default 0;
alter table public.client_aml add column if not exists ubo_change_date date;   -- change in beneficial owners noticed on
alter table public.client_aml add column if not exists ubo_filed_date date;    -- client filed the change in ΚΜΠΔ on (60-day rule)

-- Suspicious-transaction cases (Βήμα 15, άρ. 22/23/30 παρ. 1γ): every case
-- examined, reported or not, with the decision and its rationale.
create table if not exists public.aml_cases (
  id uuid primary key default gen_random_uuid(),
  client_id uuid references public.clients(id) on delete set null,
  detected_on date not null default current_date,
  detected_by uuid references public.profiles(id) on delete set null,
  description text not null,
  decision text not null default 'pending' check (decision in ('pending','reported','not_reported','abstained')),
  decision_on date,
  decided_by uuid references public.profiles(id) on delete set null,
  report_ref text,
  rationale text,
  created_at timestamptz not null default now()
);
create index if not exists aml_cases_client_idx on public.aml_cases(client_id, detected_on);

-- Requests from authorities (ΑΑΔΕ, Αρχή, other) that need an answer by a date.
create table if not exists public.aml_authority_requests (
  id uuid primary key default gen_random_uuid(),
  authority text not null,
  received_on date not null default current_date,
  due_on date,
  subject text not null,
  client_id uuid references public.clients(id) on delete set null,
  status text not null default 'open' check (status in ('open','answered')),
  answered_on date,
  note text,
  created_at timestamptz not null default now()
);

create table if not exists public.client_beneficial_owners (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  name text not null,
  afm text,
  id_number text,
  nationality text,
  residence_country text,
  ownership_pct numeric,
  control_type text,
  is_pep boolean not null default false,
  verified_on date,
  verification_doc text,
  note text,
  created_at timestamptz not null default now()
);
create index if not exists client_beneficial_owners_client_idx on public.client_beneficial_owners(client_id);

-- Tasks created from the onboarding checklist are flagged, so the app can
-- tell when onboarding is complete.
-- (column added below, after the tasks table exists)

-- Engagement agreement status (the draft itself is generated from the
-- client's data in the app; only status, version and signing date are stored).
alter table public.clients add column if not exists contract_status text not null default 'none';
alter table public.clients drop constraint if exists clients_contract_status_check;
alter table public.clients add constraint clients_contract_status_check check (contract_status in ('none','draft','sent','signed'));
alter table public.clients add column if not exists contract_signed_on date;
alter table public.clients add column if not exists contract_version int not null default 1;
alter table public.clients add column if not exists contract_special_terms text;
alter table public.clients add column if not exists contract_payment_terms text;
alter table public.clients add column if not exists contract_services_text text;        -- one bullet per line for article 2.1
alter table public.clients add column if not exists contract_payroll_base_fee numeric;    -- article 2.2
alter table public.clients add column if not exists contract_payroll_included int;
alter table public.clients add column if not exists contract_payroll_extra_fee numeric;
alter table public.clients add column if not exists contract_hourly_rate numeric;         -- article 2.3
alter table public.clients add column if not exists contract_notice_months int;           -- article 11
alter table public.clients add column if not exists contract_include_aml boolean not null default true;
alter table public.clients add column if not exists contract_start_date date;

-- ---------- annual & start-up service fees (client card → Fees) ----------
-- Switchable per client; "Post to revenue" creates a client_extra_fees row and
-- stores its id here so later edits keep the two in sync.
alter table public.clients add column if not exists annual_fee_enabled boolean not null default false;
alter table public.clients add column if not exists annual_fee numeric;
alter table public.clients add column if not exists annual_fee_month text;
alter table public.clients add column if not exists annual_fee_cc text default 'accounting';
alter table public.clients add column if not exists annual_fee_extra_id uuid;
alter table public.clients add column if not exists formation_fee_enabled boolean not null default false;
alter table public.clients add column if not exists formation_fee numeric;
alter table public.clients add column if not exists formation_fee_month text;
alter table public.clients add column if not exists formation_fee_cc text default 'other';
alter table public.clients add column if not exists formation_fee_extra_id uuid;

-- Client documents: files uploaded to Supabase Storage (bucket
-- "client-documents", private). Each row points to one stored file and is
-- tagged with a category — an AML document key, 'contract', or 'other'.
create table if not exists public.client_documents (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  category text not null default 'other',
  file_name text not null,
  file_path text not null,
  file_size bigint,
  mime_type text,
  note text,
  uploaded_by uuid references public.profiles(id) on delete set null,
  uploaded_at timestamptz not null default now()
);
create index if not exists client_documents_client_idx on public.client_documents(client_id, category);

-- Storage bucket for those files (private; admins only via policies below).
insert into storage.buckets (id, name, public)
select 'client-documents', 'client-documents', false
where not exists (select 1 from storage.buckets where id = 'client-documents');

-- Contact persons at the client.
create table if not exists public.client_contacts (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  name text not null,
  role text,
  email text,
  phone text,
  note text,
  is_primary boolean not null default false,
  created_at timestamptz not null default now()
);
create index if not exists client_contacts_client_idx on public.client_contacts(client_id);

-- Communication log / notes: call, email, meeting, note.
create table if not exists public.client_notes (
  id uuid primary key default gen_random_uuid(),
  client_id uuid not null references public.clients(id) on delete cascade,
  author_id uuid references public.profiles(id) on delete set null,
  kind text not null default 'note' check (kind in ('call','email','meeting','note')),
  note_date date not null default current_date,
  content text not null,
  created_at timestamptz not null default now()
);
create index if not exists client_notes_client_idx on public.client_notes(client_id, note_date);

-- Tasks: per client (optional) and per assignee, with due date, priority, status.
create table if not exists public.tasks (
  id uuid primary key default gen_random_uuid(),
  client_id uuid references public.clients(id) on delete set null,
  title text not null,
  description text,
  assigned_to uuid references public.profiles(id) on delete set null,
  created_by uuid references public.profiles(id) on delete set null,
  due_date date,
  priority text not null default 'normal' check (priority in ('low','normal','high')),
  status text not null default 'open' check (status in ('open','in_progress','done')),
  done_at date,
  created_at timestamptz not null default now()
);
alter table public.tasks add column if not exists kind text not null default 'general';
alter table public.tasks drop constraint if exists tasks_kind_check;
alter table public.tasks add constraint tasks_kind_check check (kind in ('general','onboarding'));
alter table public.tasks add column if not exists stage text;
create index if not exists tasks_assignee_idx on public.tasks(assigned_to, status, due_date);
create index if not exists tasks_client_idx on public.tasks(client_id);

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
alter table public.client_contacts enable row level security;
alter table public.client_notes enable row level security;
alter table public.tasks enable row level security;
alter table public.onboarding_templates enable row level security;
alter table public.client_aml enable row level security;
alter table public.client_documents enable row level security;
alter table public.aml_cases enable row level security;
alter table public.aml_authority_requests enable row level security;
alter table public.client_beneficial_owners enable row level security;
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

-- CRM: contacts readable by all signed-in, managed by admins.
drop policy if exists "client_contacts_select" on public.client_contacts;
create policy "client_contacts_select" on public.client_contacts
  for select using (auth.uid() is not null);
drop policy if exists "client_contacts_admin_write" on public.client_contacts;
create policy "client_contacts_admin_write" on public.client_contacts
  for all using (public.is_admin()) with check (public.is_admin());

-- CRM: notes — anyone signed in can read and add (as themselves); edit or
-- delete only your own, or any as admin.
drop policy if exists "client_notes_select" on public.client_notes;
create policy "client_notes_select" on public.client_notes
  for select using (auth.uid() is not null);
drop policy if exists "client_notes_insert" on public.client_notes;
create policy "client_notes_insert" on public.client_notes
  for insert with check (author_id = auth.uid());
drop policy if exists "client_notes_update" on public.client_notes;
create policy "client_notes_update" on public.client_notes
  for update using (author_id = auth.uid() or public.is_admin());
drop policy if exists "client_notes_delete" on public.client_notes;
create policy "client_notes_delete" on public.client_notes
  for delete using (author_id = auth.uid() or public.is_admin());

-- CRM: onboarding checklist — readable by all signed-in, managed by admins.
drop policy if exists "onboarding_templates_select" on public.onboarding_templates;
create policy "onboarding_templates_select" on public.onboarding_templates
  for select using (auth.uid() is not null);
drop policy if exists "onboarding_templates_admin_write" on public.onboarding_templates;
create policy "onboarding_templates_admin_write" on public.onboarding_templates
  for all using (public.is_admin()) with check (public.is_admin());

-- Client documents: admins only — table rows and the files themselves.
drop policy if exists "client_documents_admin_all" on public.client_documents;
create policy "client_documents_admin_all" on public.client_documents
  for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "client_documents_storage_select" on storage.objects;
create policy "client_documents_storage_select" on storage.objects
  for select using (bucket_id = 'client-documents' and public.is_admin());
drop policy if exists "client_documents_storage_insert" on storage.objects;
create policy "client_documents_storage_insert" on storage.objects
  for insert with check (bucket_id = 'client-documents' and public.is_admin());
drop policy if exists "client_documents_storage_update" on storage.objects;
create policy "client_documents_storage_update" on storage.objects
  for update using (bucket_id = 'client-documents' and public.is_admin());
drop policy if exists "client_documents_storage_delete" on storage.objects;
create policy "client_documents_storage_delete" on storage.objects
  for delete using (bucket_id = 'client-documents' and public.is_admin());

drop policy if exists "aml_cases_admin_all" on public.aml_cases;
create policy "aml_cases_admin_all" on public.aml_cases
  for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "aml_requests_admin_all" on public.aml_authority_requests;
create policy "aml_requests_admin_all" on public.aml_authority_requests
  for all using (public.is_admin()) with check (public.is_admin());

-- Due diligence records: admins only (sensitive).
drop policy if exists "client_aml_admin_all" on public.client_aml;
create policy "client_aml_admin_all" on public.client_aml
  for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "client_ubo_admin_all" on public.client_beneficial_owners;
create policy "client_ubo_admin_all" on public.client_beneficial_owners
  for all using (public.is_admin()) with check (public.is_admin());

-- CRM: tasks — employees see what is assigned to them or created by them;
-- admins see everything. Anyone can create (as themselves).
drop policy if exists "tasks_select" on public.tasks;
create policy "tasks_select" on public.tasks
  for select using (assigned_to = auth.uid() or created_by = auth.uid() or public.is_admin());
drop policy if exists "tasks_insert" on public.tasks;
create policy "tasks_insert" on public.tasks
  for insert with check (created_by = auth.uid());
drop policy if exists "tasks_update" on public.tasks;
create policy "tasks_update" on public.tasks
  for update using (assigned_to = auth.uid() or created_by = auth.uid() or public.is_admin());
drop policy if exists "tasks_delete" on public.tasks;
create policy "tasks_delete" on public.tasks
  for delete using (created_by = auth.uid() or public.is_admin());

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


-- ---------- other operating expenses (OPEX) ----------
-- Insurance, accounting software, bank charges, equipment depreciation, AI
-- costs, business development… Same shape as rental_expenses: recurring
-- (from_month → to_month, null = open-ended) or extra (single month, may be
-- negative). Allocated to a department, or 'general' (company-wide).
create table if not exists public.operating_expenses (
  id uuid primary key default gen_random_uuid(),
  kind text not null default 'recurring' check (kind in ('recurring','extra')),
  month char(7),
  from_month char(7),
  to_month char(7),
  category text not null check (category in ('accounting_software','other_software','database','ai_costs','insurance','car_insurance','car_depreciation','depreciation','website','emails','server_cost','business_development','bank_costs','efka_director','health_director','pos_charges','other')),
  department text not null default 'general' check (department in ('general','accounting_team','management_team','rental')),
  amount numeric not null,
  note text,
  created_at timestamptz not null default now(),
  constraint operating_expenses_amount_check check ((kind = 'recurring' and amount > 0) or (kind = 'extra' and amount <> 0)),
  constraint operating_expenses_period_check check (
    (kind = 'extra' and month is not null) or
    (kind = 'recurring' and from_month is not null and (to_month is null or to_month >= from_month))
  )
);
create index if not exists operating_expenses_idx on public.operating_expenses(kind, month, from_month);
alter table public.operating_expenses enable row level security;
drop policy if exists "operating_expenses_admin_all" on public.operating_expenses;
create policy "operating_expenses_admin_all" on public.operating_expenses
  for all using (public.is_admin()) with check (public.is_admin());

-- ---------- agreement: free-text other charges & clarifications ----------
alter table public.clients add column if not exists contract_other_fees_text text;

-- ---------- onboarding deadlines (from the engagement date) ----------
-- Engagement date set manually before the agreement is signed; once signed,
-- contract_signed_on is used instead. Due diligence + contract tasks fall due
-- on that date, company formation +3 working days, client setup +5.
alter table public.clients add column if not exists engagement_date date;

-- ---------- VAT rate per client (24% default; 0% for clients without VAT) ----------
-- Fees, revenue and profitability stay net; the Billing monitor shows net + VAT = payable.
alter table public.clients add column if not exists vat_rate numeric not null default 24;
alter table public.clients drop constraint if exists clients_vat_rate_check;
alter table public.clients add constraint clients_vat_rate_check check (vat_rate in (0, 24));

-- ---------- operating expense categories (v12) ----------
-- All existing rows keep their category; the list below is the new set (Bank charges stays).
alter table public.operating_expenses drop constraint if exists operating_expenses_category_check;
alter table public.operating_expenses add constraint operating_expenses_category_check check (category in ('accounting_software','other_software','database','ai_costs','insurance','car_insurance','car_depreciation','depreciation','website','emails','server_cost','business_development','bank_costs','efka_director','health_director','pos_charges','other'));

-- ---------------------------------------------------------------------------
-- Invoicing export (myDATA coding): invoice type, VAT exemption reason, service
-- ---------------------------------------------------------------------------
alter table public.clients add column if not exists invoice_type text not null default '2.1';
alter table public.clients add column if not exists vat_exemption_category integer;
alter table public.clients add column if not exists fixed_service text not null default 'accounting';
alter table public.clients drop constraint if exists clients_invoice_type_check;
alter table public.clients add constraint clients_invoice_type_check check (invoice_type in ('2.1', '2.2', '2.3'));
alter table public.clients drop constraint if exists clients_fixed_service_check;
alter table public.clients add constraint clients_fixed_service_check check (fixed_service in ('accounting', 'payroll', 'tax', 'formation', 'advisory', 'other'));


-- ============================================================
-- Clients & Billing open to all staff (full edit, all amounts).
-- Run once. time_entries policies are NOT touched: employees still
-- see only their own hours. Due diligence / AML / documents stay admin-only.
-- ============================================================
create or replace function public.is_staff()
returns boolean
language sql
security definer set search_path = public
stable
as $$
  select exists (select 1 from public.profiles where id = auth.uid() and active = true);
$$;

drop policy if exists "clients_insert_admin" on public.clients;
drop policy if exists "clients_insert_staff" on public.clients;
create policy "clients_insert_staff" on public.clients for insert with check (public.is_staff());
drop policy if exists "clients_update_admin" on public.clients;
drop policy if exists "clients_update_staff" on public.clients;
create policy "clients_update_staff" on public.clients for update using (public.is_staff());

drop policy if exists "fee_history_admin_all" on public.client_fee_history;
drop policy if exists "fee_history_staff_all" on public.client_fee_history;
create policy "fee_history_staff_all" on public.client_fee_history for all using (public.is_staff()) with check (public.is_staff());

drop policy if exists "extra_fees_admin_all" on public.client_extra_fees;
drop policy if exists "extra_fees_staff_all" on public.client_extra_fees;
create policy "extra_fees_staff_all" on public.client_extra_fees for all using (public.is_staff()) with check (public.is_staff());

drop policy if exists "client_contacts_admin_write" on public.client_contacts;
drop policy if exists "client_contacts_staff_write" on public.client_contacts;
create policy "client_contacts_staff_write" on public.client_contacts for all using (public.is_staff()) with check (public.is_staff());

drop policy if exists "billing_status_admin_all" on public.billing_status;
drop policy if exists "billing_status_staff_all" on public.billing_status;
create policy "billing_status_staff_all" on public.billing_status for all using (public.is_staff()) with check (public.is_staff());


-- ============================================================
-- Agreement / Due diligence / Documents of the client card open to all staff.
-- Run once, after the "Clients & Billing open to all staff" block.
-- The AML Compliance page (cases, authority requests, firm settings) stays admin-only.
-- ============================================================
drop policy if exists "client_documents_admin_all" on public.client_documents;
drop policy if exists "client_documents_staff_all" on public.client_documents;
create policy "client_documents_staff_all" on public.client_documents for all using (public.is_staff()) with check (public.is_staff());

drop policy if exists "client_documents_storage_select" on storage.objects;
create policy "client_documents_storage_select" on storage.objects for select using (bucket_id = 'client-documents' and public.is_staff());
drop policy if exists "client_documents_storage_insert" on storage.objects;
create policy "client_documents_storage_insert" on storage.objects for insert with check (bucket_id = 'client-documents' and public.is_staff());
drop policy if exists "client_documents_storage_update" on storage.objects;
create policy "client_documents_storage_update" on storage.objects for update using (bucket_id = 'client-documents' and public.is_staff());
drop policy if exists "client_documents_storage_delete" on storage.objects;
create policy "client_documents_storage_delete" on storage.objects for delete using (bucket_id = 'client-documents' and public.is_staff());

drop policy if exists "client_aml_admin_all" on public.client_aml;
drop policy if exists "client_aml_staff_all" on public.client_aml;
create policy "client_aml_staff_all" on public.client_aml for all using (public.is_staff()) with check (public.is_staff());

drop policy if exists "client_ubo_admin_all" on public.client_beneficial_owners;
drop policy if exists "client_ubo_staff_all" on public.client_beneficial_owners;
create policy "client_ubo_staff_all" on public.client_beneficial_owners for all using (public.is_staff()) with check (public.is_staff());


-- Which of our bank accounts the client pays our fee into (piraeus / revolut). Run once.
alter table public.clients add column if not exists payment_bank text;
alter table public.clients drop constraint if exists clients_payment_bank_check;
alter table public.clients add constraint clients_payment_bank_check check (payment_bank is null or payment_bank in ('piraeus','revolut'));


-- ============================================================
-- SALES PIPELINE (lead → won), manual leads, reminders via tasks.  Run once.
-- ============================================================
-- 1) Lifecycle gets a terminal "lost" state (lead that did not sign).
alter table public.clients drop constraint if exists clients_lifecycle_check;
alter table public.clients add constraint clients_lifecycle_check check (lifecycle in ('lead','onboarding','live','ended','lost'));

-- 2) Pipeline fields on the client (used while lifecycle = 'lead'; owner = account_manager_id).
alter table public.clients add column if not exists pipeline_stage text not null default 'new';
alter table public.clients drop constraint if exists clients_pipeline_stage_check;
alter table public.clients add constraint clients_pipeline_stage_check check (pipeline_stage in ('new','contact','proposal','agreement'));
alter table public.clients add column if not exists stage_changed_at timestamptz not null default now();
alter table public.clients add column if not exists lost_reason text;
alter table public.clients add column if not exists lost_at timestamptz;
alter table public.clients add column if not exists won_at timestamptz;

-- 3) Reminder bookkeeping on general tasks (snooze from Slack, one reminder per day).
alter table public.tasks add column if not exists snoozed_until date;
alter table public.tasks add column if not exists last_reminded_on date;

-- 4) Pipeline history + queue for the shared-channel notifications.
create table if not exists public.pipeline_events (
  id uuid primary key default gen_random_uuid(),
  client_id uuid references public.clients(id) on delete cascade,
  client_name text,
  event text not null check (event in ('new','stage','won','lost','reopened')),
  from_stage text,
  to_stage text,
  reason text,
  by_user uuid references public.profiles(id) on delete set null,
  notified boolean not null default false,
  created_at timestamptz not null default now()
);
create index if not exists pipeline_events_client_idx on public.pipeline_events(client_id, created_at);
create index if not exists pipeline_events_pending_idx on public.pipeline_events(notified) where notified = false;
alter table public.pipeline_events enable row level security;
drop policy if exists "pipeline_events_staff_select" on public.pipeline_events;
create policy "pipeline_events_staff_select" on public.pipeline_events for select using (public.is_staff());

-- Trigger: keeps stage_changed_at / won_at / lost_at and writes pipeline_events.
create or replace function public.track_pipeline() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if new.lifecycle = 'lead' and old.lifecycle = 'lead' and new.pipeline_stage is distinct from old.pipeline_stage then
    new.stage_changed_at := now();
    insert into public.pipeline_events(client_id, client_name, event, from_stage, to_stage, by_user, notified) values (new.id, new.name, 'stage', old.pipeline_stage, new.pipeline_stage, auth.uid(), true);
  end if;
  if old.lifecycle = 'lead' and new.lifecycle in ('onboarding','live') then
    new.won_at := coalesce(new.won_at, now());
    insert into public.pipeline_events(client_id, client_name, event, from_stage, by_user) values (new.id, new.name, 'won', old.pipeline_stage, auth.uid());
  elsif old.lifecycle = 'lead' and new.lifecycle = 'lost' then
    new.lost_at := now();
    insert into public.pipeline_events(client_id, client_name, event, from_stage, reason, by_user) values (new.id, new.name, 'lost', old.pipeline_stage, new.lost_reason, auth.uid());
  elsif old.lifecycle = 'lost' and new.lifecycle = 'lead' then
    new.lost_at := null; new.lost_reason := null; new.pipeline_stage := 'new'; new.stage_changed_at := now();
    insert into public.pipeline_events(client_id, client_name, event, to_stage, by_user, notified) values (new.id, new.name, 'reopened', 'new', auth.uid(), true);
  end if;
  return new;
end $$;
drop trigger if exists clients_track_pipeline on public.clients;
drop trigger if exists clients_track_pipeline_upd on public.clients;
create trigger clients_track_pipeline_upd before update on public.clients
  for each row execute function public.track_pipeline();
create or replace function public.track_pipeline_insert() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if new.lifecycle = 'lead' then
    insert into public.pipeline_events(client_id, client_name, event, to_stage, by_user) values (new.id, new.name, 'new', new.pipeline_stage, auth.uid());
  end if;
  return new;
end $$;
drop trigger if exists clients_track_pipeline_ins on public.clients;
create trigger clients_track_pipeline_ins after insert on public.clients
  for each row execute function public.track_pipeline_insert();


-- Lead offer (προσφορά): fee analysis + services text, used for the PDF and carried into the agreement on "Won". Run once.
alter table public.clients add column if not exists proposal jsonb;
