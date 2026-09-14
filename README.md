# HoursLedger — GLORIA (ACCOUNTING ACCO L.P.)

A real, independently-hosted internal tool for tracking employee hours per
client, monitoring billing against agreed terms, and managing time off.

- **Frontend**: React (built with Vite).
- **Backend**: [Supabase](https://supabase.com) — a hosted Postgres database
  plus real user accounts (sign up / sign in), used directly from the
  frontend. No separate server to run or maintain.

Once deployed, it has its own URL and works from any device/browser for
as many people as you invite.

---

## What's in this version

- **Timesheets**: employees log hours per client per day, with a service
  category (Accounting / Tax / Payroll / Other) and an optional note.
  Weeks can be submitted (locked) and reopened.
- **Weekly hours vs. contract**: each employee has a contracted
  hours/week figure; the manager sees who's over or under.
- **Time off**: five leave types (Annual, Sick, Medical, Other — provided
  by law, Extra — from ACCO). Annual leave has a per-employee yearly
  entitlement, and the balance **carries over** from year to year
  (so it can go negative). "Other (by law)" requires a note.
- **Clients & billing**: a fixed monthly fee plus optional one-off
  **extra fees** billed for a specific month — not tied to hours worked.
  A separate **category allocation** (Accounting/Tax/Payroll/Other hours
  per month) tracks workload only, compared against actual hours logged.
- **Full history, never overwritten**: both client fees/allocation and
  employee labor cost keep a dated history. Editing "today's" numbers
  never changes what a past month's report shows — the Dashboard always
  uses whatever was in effect during the period being viewed.
- **Dashboard**: Weekly / Monthly / Quarterly / Annual views, revenue,
  labor cost, profit and margin — both company-wide and broken down per
  client and per employee — plus a 6-month trend.
- **Calendar**: Greek public holidays (calculated, including the movable
  Orthodox Easter-linked ones) and approved time off.

---

## 1. Create the database (Supabase) — ~5 minutes

**Already have this project running?** Just re-run the updated
`supabase/schema.sql` in the SQL Editor (it's safe — every statement
checks for existing objects first). This version adds an `app_settings`
table (used by the Dashboard's "Customize" panel) and tightens security
so only a Manager can reopen a submitted week — an employee can still
lock/submit their own, but not undo it themselves.

1. Go to [supabase.com](https://supabase.com), sign up, and click **New project**.
   Pick any name and a database password (save it somewhere safe).
2. Open the **SQL Editor** (left sidebar).
3. Open `supabase/schema.sql` from this folder, copy its entire contents,
   paste it into the SQL Editor, and click **Run**. It creates every
   table, index, and security rule. It's safe to re-run if you ever need
   to (every statement checks for existing objects first).
4. Go to **Authentication → Providers** and make sure **Email** is enabled
   (it is by default).
5. Recommended for an internal tool: go to **Authentication → Settings**
   and turn **off** "Confirm email" so people can sign in immediately
   after creating an account. (Leave it on if you'd rather require email
   verification.)
6. Go to **Project Settings → API** and note down:
   - **Project URL**
   - **anon public** key (never the `service_role` key)

## 2. Connect the app to your database

Copy `.env.example` to `.env` and fill in the two values from step 1.6:
```
VITE_SUPABASE_URL=https://your-project-ref.supabase.co
VITE_SUPABASE_ANON_KEY=your-anon-public-key
```

## 3. Try it locally (optional, needs Node.js)

```bash
npm install
npm run dev
```
Open the printed URL (usually `http://localhost:5173`).

## 4. Create your first account and make it a manager

1. Open the app and click **"Need an account? Sign up"** — use your own
   name, work email, and a password. New accounts start as **Employee**.
2. In the Supabase SQL Editor, promote yourself:
   ```sql
   update public.profiles set role = 'admin' where email = 'you@acco.gr';
   ```
3. Refresh the app — you'll now see Dashboard, Clients, Employees, and
   Leave Approvals.
4. Everyone else signs up the same way; find them afterwards on the
   **Employees** page to set their role, contracted hours, leave
   entitlement, and labor cost.

## 5. Deploy it so anyone can open it from any device

The simplest option is **[Vercel](https://vercel.com)** (Netlify is
almost identical):

1. Push this project to a GitHub repository.
2. On [vercel.com](https://vercel.com), **Add New → Project**, import
   that repository. Vercel auto-detects the Vite build — leave settings
   as-is.
3. Under **Environment Variables**, add `VITE_SUPABASE_URL` and
   `VITE_SUPABASE_ANON_KEY` (same values as your `.env`).
4. Click **Deploy**. You'll get a live URL in about a minute.

Pushing a change to GitHub afterwards redeploys automatically.

---

## How roles and access work

- **Everyone signs themselves up** with a work email and password — no
  admin-set passwords. New accounts start as Employee.
- A **Manager** promotes someone to Manager, sets their contracted hours,
  annual leave entitlement, labor cost, and deactivates accounts, all
  from the **Employees** page. There's no "Add employee" button — people
  join by signing up.
- **Clients** are created directly by a Manager on the **Clients** page —
  set the category allocation, fixed fee, and effective date; add extra
  fees for specific months with the quick **+** button next to a client.
- Employees only ever see and edit **their own** hours and leave
  requests — enforced both in the interface and directly in the database
  (Postgres Row Level Security), so it holds even if the app itself were
  bypassed.
- Labor cost and fee history are visible only to Managers (Row Level
  Security restricts those tables to admins).
- "Submit week" is a workflow convenience (stops edits to a finalized
  week) rather than a hard security boundary — a manager can always
  reopen it.

## Data model notes

- **Employee cost** (`employee_cost_history`) and **client fees**
  (`client_fee_history`) are append-only: saving new numbers adds a row
  effective from a chosen date rather than overwriting the last one. The
  Dashboard picks whichever row was in effect for each day of the period
  it's showing — so a mid-month rate change is split proportionally,
  automatically.
- **Extra fees** (`client_extra_fees`) are a simple one-off ledger keyed
  by month (`YYYY-MM`) — not versioned, since each entry already applies
  to one specific month.
- **Category allocation** is for workload tracking only (allocated vs.
  actual hours by Accounting/Tax/Payroll/Other) — it does not affect
  billing.

## Notes and limitations (this is intentionally an MVP)

- No password-reset email flow is wired up yet (Supabase supports one —
  ask if you'd like it added).
- No CSV export or invoicing — by design, to keep this focused on hours,
  billing terms, and profitability.
- The free Supabase tier is more than enough for a small team; check
  Supabase's pricing page if your team or history grows significantly.
