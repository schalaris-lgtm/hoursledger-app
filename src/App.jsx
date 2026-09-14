import React, { useState, useEffect, useCallback } from "react";
import {
  LayoutDashboard, Users, Building2, Clock, Plus, Pencil, Trash2,
  ChevronLeft, ChevronRight, LogOut, X, Lock, Unlock, AlertTriangle,
  Check, Search, CalendarDays, CalendarRange, ClipboardCheck, Euro
} from "lucide-react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, Legend
} from "recharts";
import * as XLSX from "xlsx";
import { supabase, supabaseConfigured } from "./supabaseClient";

/* ---------------------------------------------------------------------- */
/* Design tokens                                                          */
/* ---------------------------------------------------------------------- */

const C = {
  bg: "#FAFAF8",
  surface: "#FFFFFF",
  ink: "#1B1E1C",
  inkMuted: "#6E7570",
  inkFaint: "#9AA19B",
  border: "#E3E1DA",
  borderStrong: "#CFCCC2",
  sidebarBg: "#12211E",
  sidebarBgActive: "#1E3A34",
  sidebarText: "#B9CCC5",
  sidebarTextActive: "#FFFFFF",
  accent: "#1C7C67",
  accentDark: "#145F4F",
  accentSoft: "#E6F1EE",
  warn: "#A86A17",
  warnSoft: "#FBF0DD",
  danger: "#AE362B",
  dangerSoft: "#FBEAE8",
};

const sans =
  "'IBM Plex Sans', ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
const mono =
  "'IBM Plex Mono', ui-monospace, 'SF Mono', 'Roboto Mono', Menlo, monospace";

const LOGO_DATA_URI = "data:image/png;base64,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";

const FontImport = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
    * { box-sizing: border-box; }
    ::selection { background: ${C.accentSoft}; }
    input:focus, select:focus, button:focus, textarea:focus {
      outline: 2px solid ${C.accent}; outline-offset: 1px;
    }
    input::placeholder { color: ${C.inkFaint}; }
    .scrollbar-thin::-webkit-scrollbar { height: 6px; width: 6px; }
    .scrollbar-thin::-webkit-scrollbar-thumb { background: ${C.borderStrong}; border-radius: 3px; }
  `}</style>
);

const TODAY = new Date();
const CATEGORIES = ["accounting", "tax", "payroll", "other"];
const CATEGORY_LABELS = { accounting: "Accounting", tax: "Tax", payroll: "Payroll", other: "Other" };

/* ---------------------------------------------------------------------- */
/* Date helpers                                                            */
/* ---------------------------------------------------------------------- */

const pad = (n) => String(n).padStart(2, "0");
const toKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromKey = (k) => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (d, n) => { const r = new Date(d); r.setDate(r.getDate() + n); return r; };
const startOfWeek = (d) => { const r = new Date(d); const day = (r.getDay() + 6) % 7; return addDays(r, -day); };
const startOfMonth = (d) => new Date(d.getFullYear(), d.getMonth(), 1);
const endOfMonth = (d) => new Date(d.getFullYear(), d.getMonth() + 1, 0);
const sameMonth = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
const fmtShort = (d) => d.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
const fmtDow = (d) => d.toLocaleDateString("en-GB", { weekday: "short" });
const fmtMonthYear = (d) => d.toLocaleDateString("en-GB", { month: "long", year: "numeric" });
const isWeekend = (d) => d.getDay() === 0 || d.getDay() === 6;
const isToday = (d) => toKey(d) === toKey(new Date());
/* ---------------------------------------------------------------------- */
/* Reporting period helpers — shared by Dashboard, Clients, Employees      */
/* ---------------------------------------------------------------------- */

const PERIOD_OPTIONS = [
  { key: "week", label: "Weekly" },
  { key: "month", label: "Monthly" },
  { key: "quarter", label: "Quarterly" },
  { key: "year", label: "Annual" },
];

function periodRange(period, anchor) {
  if (period === "week") { const s = startOfWeek(anchor); return [s, addDays(s, 6)]; }
  if (period === "quarter") {
    const q = Math.floor(anchor.getMonth() / 3);
    return [new Date(anchor.getFullYear(), q * 3, 1), new Date(anchor.getFullYear(), q * 3 + 3, 0)];
  }
  if (period === "year") return [new Date(anchor.getFullYear(), 0, 1), new Date(anchor.getFullYear(), 11, 31)];
  return [startOfMonth(anchor), endOfMonth(anchor)]; // "month" default
}
function shiftPeriodAnchor(period, anchor, dir) {
  if (period === "week") return addDays(anchor, dir * 7);
  if (period === "quarter") return new Date(anchor.getFullYear(), anchor.getMonth() + dir * 3, 1);
  if (period === "year") return new Date(anchor.getFullYear() + dir, anchor.getMonth(), 1);
  return new Date(anchor.getFullYear(), anchor.getMonth() + dir, 1); // "month"
}
function periodLabel(period, rangeStart, rangeEnd) {
  if (period === "week") return `${fmtShort(rangeStart)} – ${fmtShort(rangeEnd)}`;
  if (period === "quarter") return `Q${Math.floor(rangeStart.getMonth() / 3) + 1} ${rangeStart.getFullYear()}`;
  if (period === "year") return `${rangeStart.getFullYear()}`;
  return fmtMonthYear(rangeStart); // "month"
}
const LEAVE_TYPES = {
  annual: "Annual leave",
  sick: "Sick leave",
  medical: "Medical leave",
  other_legal: "Other (provided by law)",
  extra_acco: "Extra (from ACCO)",
};
// Types that require a note explaining the reason, since they cover
// statutory leave categories (marriage, bereavement, etc.).
const LEAVE_TYPES_REQUIRING_NOTE = new Set(["other_legal"]);
// Older data (from before these categories existed) used different keys —
// keep counting and labeling those correctly instead of showing "0"/blank.
const LEGACY_LEAVE_TYPE_MAP = { vacation: "annual", other: "other_legal" };
function normalizedLeaveType(type) { return LEGACY_LEAVE_TYPE_MAP[type] || type; }
function leaveTypeLabel(type) { return LEAVE_TYPES[normalizedLeaveType(type)] || type; }

function leaveStatusTone(status) {
  return status === "approved" ? "accent" : status === "rejected" ? "danger" : "warn";
}
/* ---------------------------------------------------------------------- */
/* Client revenue / productivity helpers                                   */
/* ---------------------------------------------------------------------- */

// Revenue a client generates for a given number of hours in a period.
// - Fixed fee (monthly retainer) is prorated for a weekly view.
// - Hourly rate applies to overage hours beyond the agreed allocation
//   when both a fixed fee and an allocation are set (retainer + overage
//   billing); otherwise it applies to all hours worked (pure hourly billing).
// Look up the historical record in effect on a given date: the latest
// entry whose effectiveDate is on or before dateKey. Falls back to the
// earliest entry if dateKey predates all history, and to `fallback` if
// there's no history at all.
function valueAsOf(history, dateKey, fallback) {
  if (!history || history.length === 0) return fallback;
  const sorted = [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  let result = sorted[0];
  for (const h of sorted) {
    if (h.effectiveDate <= dateKey) result = h;
    else break;
  }
  return result;
}

// The earliest effective date recorded in a fee/cost history — used as
// "when did this client/employee's terms actually start applying",
// which the admin controls explicitly (e.g. backdating to when a
// long-standing client was first entered into the system), rather than
// whenever the database row happened to be created.
function earliestHistoryDate(history) {
  if (!history || history.length === 0) return null;
  return [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate))[0].effectiveDate;
}

// Sorts a history array chronologically and adds a derived "from"/"to"
// period to each entry — "to" is the day before the next entry starts,
// or "Present" for the most recent one. Purely for display.
function withPeriods(history) {
  const sorted = [...(history || [])].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  return sorted.map((h, i) => ({
    ...h,
    from: h.effectiveDate,
    to: i < sorted.length - 1 ? toKey(addDays(fromKey(sorted[i + 1].effectiveDate), -1)) : "Present",
  }));
}

// After editing or deleting a history entry, the employee/client's
// top-level "current" fields should track whichever entry is now the
// most recent by effective date.
function latestFromHistory(history, fieldNames) {
  if (!history || history.length === 0) return null;
  const sorted = [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  const latest = sorted[sorted.length - 1];
  const result = {};
  fieldNames.forEach((f) => { result[f] = latest[f]; });
  return result;
}

function currentFeeOf(client) {
  return {
    fixedFee: client.fixedFee,
    allocation: client.allocation || { accounting: 0, tax: 0, payroll: 0, other: 0 },
  };
}
// The fee terms (fixed fee / category allocation) in effect on a given
// date, e.g. the start of the period being reported on. Extra fees are
// tracked separately as one-off monthly entries — see extraFeeEntries.
function feeAsOf(client, dateKey) {
  return valueAsOf(client.feeHistory, dateKey, currentFeeOf(client));
}

/* ---------------------------------------------------------------------- */
/* Precise period math: splits a date range at every point the fee/cost    */
/* actually changed, and prorates each segment by its share of days —      */
/* so a mid-period change is reflected exactly, not just at period start.  */
/* ---------------------------------------------------------------------- */

const daysBetweenInclusive = (a, b) => Math.round((b - a) / 86400000) + 1;

// Splits [rangeStart, rangeEnd] into contiguous segments, one per distinct
// historical value in effect, using effectiveDate breakpoints that fall
// inside the range.
function regimesInRange(history, fallbackValue, rangeStart, rangeEnd) {
  const rangeStartKey = toKey(rangeStart);
  const rangeEndKey = toKey(rangeEnd);
  const sorted = history && history.length ? [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate)) : null;
  const breakpoints = sorted
    ? [...new Set(sorted.map((h) => h.effectiveDate).filter((d) => d > rangeStartKey && d <= rangeEndKey))].sort()
    : [];
  const boundaries = [rangeStartKey, ...breakpoints];
  return boundaries.map((segStartKey, i) => {
    const segEndKey = i + 1 < boundaries.length ? toKey(addDays(fromKey(boundaries[i + 1]), -1)) : rangeEndKey;
    return { start: fromKey(segStartKey), end: fromKey(segEndKey), value: valueAsOf(history, segStartKey, fallbackValue) };
  });
}

// Revenue AND allocation for a client over an arbitrary range, correctly
// prorated across any fee changes that happened mid-range.
// Revenue AND allocation for a client over an ARBITRARY date range (a
// week, month, quarter, year, whatever) — accrued day by day, using each
// day's own calendar-month length. This makes the math period-agnostic
// and exact even when a fee changed mid-range.
// Revenue, and category-by-category allocation vs. actual hours, for a
// client over an ARBITRARY date range (week, month, quarter, year, ...).
// Billing is a flat fixed fee + extra fees (no more hourly overage), each
// accrued day by day using each day's own calendar-month length so a
// mid-range fee change is reflected exactly and the math is period-agnostic.
// Sum of a client's one-off monthly extra fees that fall within a date
// range, accrued day by day (each entry's amount is spread evenly across
// the days of the month it was billed for) — so it works for any period
// length and any mix of months.
function extraFeesAccruedForRange(client, rangeStart, rangeEnd) {
  const endDate = client.endDate ? fromKey(client.endDate) : null;
  let total = 0;
  let d = new Date(rangeStart);
  while (d <= rangeEnd) {
    if (!endDate || d <= endDate) {
      const monthKey = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
      const monthlyTotal = (client.extraFeeEntries || [])
        .filter((e) => e.month === monthKey)
        .reduce((s, e) => s + (Number(e.amount) || 0), 0);
      if (monthlyTotal) total += monthlyTotal / endOfMonth(d).getDate();
    }
    d = addDays(d, 1);
  }
  return total;
}

function periodClientMetrics(client, periodEntries, rangeStart, rangeEnd) {
  const startDateKey = earliestHistoryDate(client.feeHistory);
  const startDate = startDateKey ? fromKey(startDateKey) : null;
  const endDate = client.endDate ? fromKey(client.endDate) : null;
  const regimes = regimesInRange(client.feeHistory, currentFeeOf(client), rangeStart, rangeEnd);
  let revenue = extraFeesAccruedForRange(client, rangeStart, rangeEnd);
  const allocationByCategory = { accounting: 0, tax: 0, payroll: 0, other: 0 };
  let hasAllocation = false;
  regimes.forEach((seg) => {
    const fee = seg.value;
    const alloc = fee.allocation || {};
    let d = new Date(seg.start);
    while (d <= seg.end) {
      // Skip any day before the fee terms you set actually started
      // applying — controlled by the "Effective from" date you choose —
      // and any day after the client's end date, if one is set.
      if ((!startDate || d >= startDate) && (!endDate || d <= endDate)) {
        const daysInThisMonth = endOfMonth(d).getDate();
        if (fee.fixedFee) revenue += fee.fixedFee / daysInThisMonth;
        CATEGORIES.forEach((cat) => {
          if (alloc[cat]) { allocationByCategory[cat] += alloc[cat] / daysInThisMonth; hasAllocation = true; }
        });
      }
      d = addDays(d, 1);
    }
  });
  const actualByCategory = { accounting: 0, tax: 0, payroll: 0, other: 0 };
  periodEntries.forEach((e) => {
    const cat = CATEGORIES.includes(e.category) ? e.category : "other";
    actualByCategory[cat] += e.hours;
  });
  const allocation = hasAllocation ? CATEGORIES.reduce((s, c) => s + allocationByCategory[c], 0) : null;
  return { revenue, allocation, allocationByCategory, actualByCategory };
}

// Labor cost for an employee over an arbitrary date range, accrued day by
// day the same way — correct across any salary/benefit change mid-range,
// and for any period length (week, month, quarter, year).
function periodEmployeeCost(employee, rangeStart, rangeEnd) {
  const regimes = regimesInRange(employee.costHistory, currentCostOf(employee), rangeStart, rangeEnd);
  let cost = 0;
  regimes.forEach((seg) => {
    const c = seg.value;
    const monthlyFull = (c.grossSalary || 0) + (c.socialSecurity || 0) + (c.ticketRestaurant || 0) + (c.insurance || 0) + (c.otherCost || 0);
    let d = new Date(seg.start);
    while (d <= seg.end) {
      cost += monthlyFull / endOfMonth(d).getDate();
      d = addDays(d, 1);
    }
  });
  return cost;
}
const fmtEur = (n) => "€" + n.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
/* ---------------------------------------------------------------------- */
/* Greek public holidays & annual leave balance helpers                    */
/* ---------------------------------------------------------------------- */

const ANNUAL_LEAVE_DAYS = 25;
const LEAVE_TRACKING_START_YEAR = 2026; // balances carry over starting from this year

// Orthodox Easter (Gregorian date) via the Meeus Julian algorithm,
// converted from the Julian to the Gregorian calendar (+13 days,
// accurate for 1900–2099, which comfortably covers this app's use).
function orthodoxEaster(year) {
  const a = year % 4;
  const b = year % 7;
  const c = year % 19;
  const d = (19 * c + 15) % 30;
  const e = (2 * a + 4 * b - d + 34) % 7;
  const month = Math.floor((d + e + 114) / 31); // 3 = March, 4 = April (Julian)
  const day = ((d + e + 114) % 31) + 1;
  const julian = new Date(year, month - 1, day);
  return addDays(julian, 13);
}

// Fixed-date and Easter-linked Greek public holidays for a given year.
function getGreekPublicHolidays(year) {
  const easter = orthodoxEaster(year);
  const list = [
    { date: new Date(year, 0, 1), name: "New Year's Day" },
    { date: new Date(year, 0, 6), name: "Epiphany" },
    { date: addDays(easter, -48), name: "Clean Monday" },
    { date: new Date(year, 2, 25), name: "Independence Day" },
    { date: addDays(easter, -2), name: "Good Friday" },
    { date: addDays(easter, 1), name: "Easter Monday" },
    { date: new Date(year, 4, 1), name: "Labour Day" },
    { date: addDays(easter, 50), name: "Whit Monday" },
    { date: new Date(year, 7, 15), name: "Assumption Day" },
    { date: new Date(year, 9, 28), name: "Ohi Day" },
    { date: new Date(year, 11, 25), name: "Christmas Day" },
    { date: new Date(year, 11, 26), name: "Glorifying Mother of God" },
  ];
  return list;
}

// Holiday lookup map (dateKey -> name) spanning the given years.
function holidayMapForYears(years) {
  const map = new Map();
  years.forEach((y) => getGreekPublicHolidays(y).forEach((h) => map.set(toKey(h.date), h.name)));
  return map;
}

// Working days between two dates (inclusive), excluding weekends and
// the given set of public-holiday date keys.
function countWorkingDays(startKey, endKey, holidayMap) {
  let d = fromKey(startKey);
  const end = fromKey(endKey);
  let count = 0;
  while (d <= end) {
    if (!isWeekend(d) && !holidayMap.has(toKey(d))) count++;
    d = addDays(d, 1);
  }
  return count;
}

// Annual-leave days already used (approved, working days only) by an
// employee within a calendar year. Only the "annual" leave type counts
// against the entitlement — medical/other/extra do not.
function vacationDaysUsed(employeeId, year, leaveRequests, holidayMap) {
  return leaveRequests
    .filter((r) => r.employeeId === employeeId && normalizedLeaveType(r.type) === "annual" && r.status === "approved" && fromKey(r.startDate).getFullYear() === year)
    .reduce((sum, r) => sum + countWorkingDays(r.startDate, r.endDate, holidayMap), 0);
}

// Approved days used within a year, broken down by leave type — e.g.
// { annual: 12, sick: 2, medical: 0, other_legal: 1, extra_acco: 0 }.
function daysUsedByType(employeeId, year, leaveRequests, holidayMap) {
  const totals = {};
  Object.keys(LEAVE_TYPES).forEach((t) => { totals[t] = 0; });
  leaveRequests
    .filter((r) => r.employeeId === employeeId && r.status === "approved" && fromKey(r.startDate).getFullYear() === year)
    .forEach((r) => {
      const t = normalizedLeaveType(r.type);
      totals[t] = (totals[t] || 0) + countWorkingDays(r.startDate, r.endDate, holidayMap);
    });
  return totals;
}

// Cumulative annual-leave balance through the end of `uptoYear` — unused
// days (or a deficit) carry forward from every earlier tracked year.
// Positive = days still owed to the employee; negative = taken in advance.
function leaveBalanceThroughYear(employeeId, entitlement, uptoYear, leaveRequests) {
  let balance = 0;
  for (let y = LEAVE_TRACKING_START_YEAR; y <= uptoYear; y++) {
    const holidayMap = holidayMapForYears([y - 1, y, y + 1]);
    const used = vacationDaysUsed(employeeId, y, leaveRequests, holidayMap);
    balance += entitlement - used;
  }
  return balance;
}
/* ---------------------------------------------------------------------- */
/* Labor cost & productivity helpers                                       */
/* ---------------------------------------------------------------------- */

const MONTHLY_WORK_DAYS = 22; // standard basis for converting a monthly cost into an hourly rate

// Total monthly employer cost: gross salary + benefits + other.
function monthlyLaborCost(e) {
  return (e.grossSalary || 0) + (e.socialSecurity || 0) + (e.ticketRestaurant || 0) + (e.insurance || 0) + (e.otherCost || 0);
}

// Hourly cost, using a standard 22-working-day month and the employee's
// contracted weekly hours (converted to a daily figure over a 5-day week).
function hourlyLaborCost(e) {
  const dailyHours = (e.weeklyHours || 40) / 5;
  const monthlyHoursBasis = MONTHLY_WORK_DAYS * dailyHours;
  return monthlyHoursBasis > 0 ? monthlyLaborCost(e) / monthlyHoursBasis : 0;
}

function currentCostOf(e) {
  return { grossSalary: e.grossSalary || 0, socialSecurity: e.socialSecurity || 0, ticketRestaurant: e.ticketRestaurant || 0, insurance: e.insurance || 0, otherCost: e.otherCost || 0 };
}
// The cost breakdown in effect on a given date (e.g. the start of the
// period being reported on), for period-accurate Dashboard figures.
function costAsOf(e, dateKey) {
  return valueAsOf(e.costHistory, dateKey, currentCostOf(e));
}
function monthlyLaborCostAsOf(e, dateKey) {
  const c = costAsOf(e, dateKey);
  return (c.grossSalary || 0) + (c.socialSecurity || 0) + (c.ticketRestaurant || 0) + (c.insurance || 0) + (c.otherCost || 0);
}
function hourlyLaborCostAsOf(e, dateKey) {
  const dailyHours = (e.weeklyHours || 40) / 5;
  const basis = MONTHLY_WORK_DAYS * dailyHours;
  return basis > 0 ? monthlyLaborCostAsOf(e, dateKey) / basis : 0;
}

/* ---------------------------------------------------------------------- */
/* Small UI atoms                                                          */
/* ---------------------------------------------------------------------- */

function Btn({ children, onClick, variant = "primary", size = "md", icon: Icon, disabled, type = "button" }) {
  const base = {
    fontFamily: sans,
    fontWeight: 600,
    borderRadius: 6,
    cursor: disabled ? "not-allowed" : "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    border: "1px solid transparent",
    transition: "background 120ms ease, border-color 120ms ease",
    opacity: disabled ? 0.5 : 1,
  };
  const sizes = { sm: { padding: "5px 10px", fontSize: 12.5 }, md: { padding: "8px 14px", fontSize: 13.5 } };
  const variants = {
    primary: { background: C.accent, color: "#fff" },
    secondary: { background: C.surface, color: C.ink, border: `1px solid ${C.border}` },
    ghost: { background: "transparent", color: C.inkMuted },
    danger: { background: C.dangerSoft, color: C.danger },
  };
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      style={{ ...base, ...sizes[size], ...variants[variant] }}
    >
      {Icon && <Icon size={14} />}
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12.5, color: C.inkMuted, fontFamily: sans }}>
      {label}
      {children}
    </label>
  );
}

const inputStyle = {
  fontFamily: sans,
  fontSize: 13.5,
  padding: "7px 9px",
  borderRadius: 6,
  border: `1px solid ${C.border}`,
  color: C.ink,
  background: C.surface,
};

function Badge({ children, tone = "neutral" }) {
  const tones = {
    neutral: { bg: C.border, fg: C.inkMuted },
    accent: { bg: C.accentSoft, fg: C.accentDark },
    warn: { bg: C.warnSoft, fg: C.warn },
    danger: { bg: C.dangerSoft, fg: C.danger },
  };
  const t = tones[tone];
  return (
    <span style={{
      fontFamily: sans, fontSize: 11, fontWeight: 600, padding: "2px 7px",
      borderRadius: 999, background: t.bg, color: t.fg, whiteSpace: "nowrap",
    }}>
      {children}
    </span>
  );
}

function Panel({ children, style }) {
  return (
    <div style={{
      background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10,
      padding: 18, ...style,
    }}>
      {children}
    </div>
  );
}

function StatCard({ label, value, sub, accent }) {
  return (
    <Panel style={{ flex: 1, minWidth: 150 }}>
      <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 8 }}>{label}</div>
      <div style={{ fontFamily: mono, fontSize: 26, fontWeight: 600, color: accent || C.ink }}>{value}</div>
      {sub && <div style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint, marginTop: 4 }}>{sub}</div>}
    </Panel>
  );
}

function Modal({ title, onClose, children, width = 420 }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(18,33,30,0.35)", zIndex: 50,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 16,
    }} onClick={onClose}>
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: C.surface, borderRadius: 12, width, maxWidth: "100%",
          padding: 22, border: `1px solid ${C.border}`, boxShadow: "0 12px 32px rgba(18,33,30,0.18)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontFamily: sans, fontSize: 16, fontWeight: 700, color: C.ink }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: C.inkMuted, padding: 4 }}>
            <X size={18} />
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

function EmptyState({ icon: Icon, title, sub }) {
  return (
    <div style={{ textAlign: "center", padding: "40px 20px", color: C.inkMuted, fontFamily: sans }}>
      <Icon size={26} style={{ marginBottom: 10, color: C.inkFaint }} />
      <div style={{ fontWeight: 600, color: C.ink, marginBottom: 4 }}>{title}</div>
      {sub && <div style={{ fontSize: 13 }}>{sub}</div>}
    </div>
  );
}


function CenteredScreen({ children }) {
  return (
    <div style={{ minHeight: "100vh", background: C.sidebarBg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: sans, padding: 20 }}>
      <FontImport />
      {children}
    </div>
  );
}

function SetupScreen() {
  return (
    <CenteredScreen>
      <div style={{ width: 460, maxWidth: "100%", background: C.surface, borderRadius: 12, padding: 26, border: `1px solid ${C.border}` }}>
        <h2 style={{ marginTop: 0, fontFamily: sans, color: C.ink }}>Almost there</h2>
        <p style={{ fontFamily: sans, fontSize: 13.5, color: C.inkMuted, lineHeight: 1.6 }}>
          This app needs a Supabase project to store its data. Copy <code>.env.example</code> to{" "}
          <code>.env</code>, fill in your project URL and anon key, then restart the dev server
          (or redeploy). See <code>README.md</code> for the full setup steps.
        </p>
      </div>
    </CenteredScreen>
  );
}

function LoadingScreen({ text = "Loading…" }) {
  return <CenteredScreen><span style={{ color: "#fff", fontFamily: sans }}>{text}</span></CenteredScreen>;
}

function AuthScreen() {
  const [mode, setMode] = useState("signin");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit() {
    if (!email.trim() || !password) { setError("Enter both your email and password."); return; }
    setError(""); setInfo(""); setBusy(true);
    if (mode === "signup") {
      const { data, error } = await supabase.auth.signUp({ email, password, options: { data: { name } } });
      if (error) setError(error.message);
      else if (!data.session) setInfo("Account created. Check your email to confirm it, then sign in.");
    } else {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) setError(error.message);
    }
    setBusy(false);
  }
  function handleKeyDown(e) {
    if (e.key === "Enter") submit();
  }

  return (
    <CenteredScreen>
      <div style={{ width: 380, maxWidth: "100%" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, marginBottom: 24 }}>
          <img src={LOGO_DATA_URI} alt="ACCO logo" style={{ width: 56, height: 56, objectFit: "contain" }} />
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 20, letterSpacing: 0.5 }}>GLORIA</span>
          <span style={{ color: C.sidebarText, fontWeight: 500, fontSize: 12 }}>ACCOUNTING ACCO L.P.</span>
        </div>
        <div style={{ background: C.surface, borderRadius: 12, padding: 22, border: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 15, color: C.ink }}>
            {mode === "signin" ? "Sign in" : "Create your account"}
          </div>
          {mode === "signup" && (
            <Field label="Full name">
              <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Maria Ioannou" />
            </Field>
          )}
          <Field label="Work email">
            <input type="text" style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={handleKeyDown} placeholder="name@acco.gr" />
          </Field>
          <Field label="Password">
            <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={handleKeyDown} placeholder="At least 6 characters" />
          </Field>
          {error && <div style={{ fontFamily: sans, fontSize: 12.5, color: C.danger, background: C.dangerSoft, padding: "8px 10px", borderRadius: 6 }}>{error}</div>}
          {info && <div style={{ fontFamily: sans, fontSize: 12.5, color: C.accentDark, background: C.accentSoft, padding: "8px 10px", borderRadius: 6 }}>{info}</div>}
          <Btn onClick={submit} disabled={busy}>{mode === "signin" ? "Sign in" : "Create account"}</Btn>
          <button type="button" onClick={() => { setMode(mode === "signin" ? "signup" : "signin"); setError(""); setInfo(""); }}
            style={{ background: "none", border: "none", cursor: "pointer", color: C.inkMuted, fontFamily: sans, fontSize: 12.5, textAlign: "center" }}>
            {mode === "signin" ? "Need an account? Sign up" : "Already have an account? Sign in"}
          </button>
        </div>
      </div>
    </CenteredScreen>
  );
}

/* ---------------------------------------------------------------------- */
/* Week strip control (shared)                                             */
/* ---------------------------------------------------------------------- */

function WeekNav({ weekStart, setWeekStart }) {
  const weekEnd = addDays(weekStart, 6);
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
      <button onClick={() => setWeekStart(addDays(weekStart, -7))} style={navBtnStyle}>
        <ChevronLeft size={16} />
      </button>
      <div style={{ fontFamily: sans, fontSize: 13.5, fontWeight: 600, color: C.ink, minWidth: 175, textAlign: "center" }}>
        {fmtShort(weekStart)} – {fmtShort(weekEnd)}, {weekEnd.getFullYear()}
      </div>
      <button onClick={() => setWeekStart(addDays(weekStart, 7))} style={navBtnStyle}>
        <ChevronRight size={16} />
      </button>
      <Btn variant="ghost" size="sm" onClick={() => setWeekStart(startOfWeek(new Date()))}>This week</Btn>
    </div>
  );
}
const navBtnStyle = {
  border: `1px solid ${C.border}`, background: C.surface, borderRadius: 6,
  width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center",
  cursor: "pointer", color: C.ink,
};

// Shared week/month/quarter/year selector — used on Clients and Employees
// so their figures reflect a chosen reporting period rather than always
// "the current month".
function PeriodSelector({ period, setPeriod, anchor, setAnchor }) {
  const [rangeStart, rangeEnd] = periodRange(period, anchor);
  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
      <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
        {PERIOD_OPTIONS.map((p) => (
          <button key={p.key} onClick={() => setPeriod(p.key)}
            style={{
              padding: "6px 10px", fontSize: 12, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer",
              background: period === p.key ? C.accent : C.surface, color: period === p.key ? "#fff" : C.inkMuted,
            }}>
            {p.label}
          </button>
        ))}
      </div>
      <button onClick={() => setAnchor((a) => shiftPeriodAnchor(period, a, -1))} style={navBtnStyle}><ChevronLeft size={16} /></button>
      <div style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, minWidth: 140, textAlign: "center" }}>
        {periodLabel(period, rangeStart, rangeEnd)}
      </div>
      <button onClick={() => setAnchor((a) => shiftPeriodAnchor(period, a, 1))} style={navBtnStyle}><ChevronRight size={16} /></button>
    </div>
  );
}

const iconBtnStyle = { background: "none", border: "none", cursor: "pointer", padding: 3, display: "flex" };

function PageHeader({ title, sub, right }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
      <div>
        <h1 style={{ fontFamily: sans, fontSize: 21, fontWeight: 700, color: C.ink, margin: 0 }}>{title}</h1>
        {sub && <div style={{ fontFamily: sans, fontSize: 13, color: C.inkMuted, marginTop: 4 }}>{sub}</div>}
      </div>
      {right}
    </div>
  );
}


function TableShell({ headers, children }) {
  return (
    <div className="scrollbar-thin" style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 13 }}>
        <thead>
          <tr>
            {headers.map((h) => (
              <th key={h} style={{
                textAlign: "left", padding: "6px 8px", fontSize: 11.5, fontWeight: 600,
                color: C.inkFaint, borderBottom: `1px solid ${C.border}`,
              }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}
function Td({ children, mono: isMono, right, style }) {
  return (
    <td style={{
      padding: "8px 8px", borderBottom: `1px solid ${C.border}`, color: C.ink,
      fontFamily: isMono ? mono : sans, textAlign: right ? "right" : "left", fontWeight: isMono ? 600 : 400,
      ...style,
    }}>{children}</td>
  );
}


function Sidebar({ user, view, setView }) {
  const employeeNav = [
    { key: "timesheet", label: "My Timesheet", icon: Clock },
    { key: "timeoff", label: "Time Off", icon: CalendarDays },
    { key: "calendar", label: "Calendar", icon: CalendarRange },
  ];
  const adminNav = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "timesheet", label: "My Timesheet", icon: Clock },
    { key: "timesheets", label: "All Timesheets", icon: Clock },
    { key: "timeoff", label: "Time Off", icon: CalendarDays },
    { key: "leaveapprovals", label: "Leave Approvals", icon: ClipboardCheck },
    { key: "calendar", label: "Calendar", icon: CalendarRange },
    { key: "clients", label: "Clients", icon: Building2 },
    { key: "employees", label: "Employees", icon: Users },
  ];
  const items = user.role === "admin" ? adminNav : employeeNav;
  return (
    <div style={{ width: 224, minWidth: 224, background: C.sidebarBg, minHeight: "100vh", display: "flex", flexDirection: "column", padding: "20px 14px", fontFamily: sans }}>
      <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 6px", marginBottom: 26 }}>
        <img src={LOGO_DATA_URI} alt="ACCO logo" style={{ width: 28, height: 28, objectFit: "contain" }} />
        <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.15 }}>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 15.5 }}>GLORIA</span>
          <span style={{ color: C.sidebarText, fontWeight: 500, fontSize: 9.5 }}>ACCOUNTING ACCO L.P.</span>
        </div>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 3, flex: 1 }}>
        {items.map((it) => {
          const active = view === it.key;
          return (
            <button key={it.key} onClick={() => setView(it.key)}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: "9px 11px", borderRadius: 7,
                border: "none", cursor: "pointer", textAlign: "left",
                background: active ? C.sidebarBgActive : "transparent",
                color: active ? C.sidebarTextActive : C.sidebarText,
                fontSize: 13.5, fontWeight: active ? 600 : 500,
              }}>
              <it.icon size={16} />
              {it.label}
            </button>
          );
        })}
      </div>
      <div style={{ borderTop: `1px solid rgba(255,255,255,0.08)`, paddingTop: 14, marginTop: 10 }}>
        <div style={{ color: "#fff", fontSize: 13, fontWeight: 600 }}>{user.name}</div>
        <div style={{ color: C.sidebarText, fontSize: 11.5, marginBottom: 10 }}>
          {user.title ? `${user.title} · ${user.role === "admin" ? "Manager" : "Employee"}` : (user.role === "admin" ? "Manager" : "Employee")}
        </div>
        <button onClick={() => supabase.auth.signOut()}
          style={{ display: "flex", alignItems: "center", gap: 7, background: "transparent", border: `1px solid rgba(255,255,255,0.12)`, color: C.sidebarText, borderRadius: 7, padding: "7px 10px", fontSize: 12.5, cursor: "pointer", width: "100%" }}>
          <LogOut size={13} /> Sign out
        </button>
      </div>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Supabase table hook: fetches once, then stays live via realtime         */
/* ---------------------------------------------------------------------- */

function useTable(table, mapRow, orderColumn) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);

  const refetch = useCallback(async () => {
    if (!supabase) { setLoading(false); return; }
    let q = supabase.from(table).select("*");
    if (orderColumn) q = q.order(orderColumn, { ascending: true });
    const { data, error } = await q;
    if (!error && data) setRows(data.map(mapRow));
    setLoading(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table]);

  useEffect(() => {
    refetch();
    if (!supabase) return;
    const channel = supabase
      .channel(`${table}_sync_${Math.random().toString(36).slice(2)}`)
      .on("postgres_changes", { event: "*", schema: "public", table }, () => refetch())
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refetch]);

  return [rows, refetch, loading];
}

const mapProfile = (r) => ({ id: r.id, name: r.name, email: r.email, title: r.title, role: r.role, weeklyHours: Number(r.weekly_hours), annualLeaveDays: Number(r.annual_leave_days), active: r.active, createdAt: r.created_at ? r.created_at.slice(0, 10) : null });
const mapCostHistory = (r) => ({ id: r.id, employeeId: r.employee_id, effectiveDate: r.effective_date, grossSalary: Number(r.gross_salary), socialSecurity: Number(r.social_security), ticketRestaurant: Number(r.ticket_restaurant), insurance: Number(r.insurance), otherCost: Number(r.other_cost) });
const mapClientRow = (r) => ({ id: r.id, name: r.name, fixedFee: r.fixed_fee === null ? null : Number(r.fixed_fee), endDate: r.end_date || null, active: r.active, createdAt: r.created_at ? r.created_at.slice(0, 10) : null });
const mapFeeHistory = (r) => ({ id: r.id, clientId: r.client_id, effectiveDate: r.effective_date, fixedFee: r.fixed_fee === null ? null : Number(r.fixed_fee), allocation: { accounting: Number(r.alloc_accounting), tax: Number(r.alloc_tax), payroll: Number(r.alloc_payroll), other: Number(r.alloc_other) } });
const mapExtraFee = (r) => ({ id: r.id, clientId: r.client_id, month: r.month, amount: Number(r.amount), note: r.note || "" });
const mapEntry = (r) => ({ id: r.id, employeeId: r.employee_id, clientId: r.client_id, date: r.entry_date, hours: Number(r.hours), category: r.category, note: r.note || "" });
const mapLockedWeek = (r) => `${r.employee_id}|${r.week_start}`;
const mapLeaveRequest = (r) => ({ id: r.id, employeeId: r.employee_id, startDate: r.start_date, endDate: r.end_date, type: r.type, note: r.note || "", status: r.status, decidedBy: r.decided_by, decidedAt: r.decided_at });

// A small shared key/value setting (e.g. Dashboard "Customize" choices),
// stored in app_settings and kept live the same way the tables above are.
function useAppSetting(key, defaultValue) {
  const [value, setValue] = useState(defaultValue);
  const [loaded, setLoaded] = useState(false);

  const refetch = useCallback(async () => {
    if (!supabase) { setLoaded(true); return; }
    const { data, error } = await supabase.from("app_settings").select("*").eq("key", key).maybeSingle();
    if (!error) setValue(data ? data.value : defaultValue);
    setLoaded(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key]);

  useEffect(() => {
    refetch();
    if (!supabase) return;
    const channel = supabase
      .channel(`app_settings_${key}_${Math.random().toString(36).slice(2)}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "app_settings", filter: `key=eq.${key}` }, () => refetch())
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refetch]);

  async function save(next) {
    setValue(next); // optimistic
    if (!supabase) return;
    await supabase.from("app_settings").upsert({ key, value: next, updated_at: new Date().toISOString() });
  }

  return [value, save, loaded];
}

// Builds an .xlsx workbook from one or more sheets and downloads it.
function exportToExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{}]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}


/* ---------------------------------------------------------------------- */
/* Employee Timesheet view                                                 */
/* ---------------------------------------------------------------------- */

function EmployeeTimesheet({ user, clients, entries, refetchEntries, lockedWeeks, refetchLockedWeeks }) {
  const [weekStart, setWeekStart] = useState(startOfWeek(TODAY));
  const [form, setForm] = useState({ date: toKey(TODAY), clientId: clients.find((c) => c.active)?.id || "", category: "accounting", hours: "", note: "" });
  const [editingId, setEditingId] = useState(null);
  const [editHours, setEditHours] = useState("");
  const [editNote, setEditNote] = useState("");

  const weekKey = `${user.id}|${toKey(weekStart)}`;
  const isLocked = lockedWeeks.includes(weekKey);
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));

  const myEntries = entries.filter((e) => e.employeeId === user.id);
  const weekEntries = myEntries.filter((e) => {
    const d = fromKey(e.date);
    return d >= weekStart && d <= addDays(weekStart, 6);
  });
  const monthEntries = myEntries.filter((e) => sameMonth(fromKey(e.date), weekStart));

  const weekTotal = weekEntries.reduce((s, e) => s + e.hours, 0);
  const monthTotal = monthEntries.reduce((s, e) => s + e.hours, 0);
  const weeklyTarget = user.weeklyHours || 40;
  const weeklyDiff = weekTotal - weeklyTarget;
  const onTargetWeek = Math.abs(weeklyDiff) < 0.01;
  const weeklyDiffLabel = onTargetWeek ? "On track" : weeklyDiff > 0 ? `+${weeklyDiff.toFixed(2)}h overtime` : `${weeklyDiff.toFixed(2)}h undertime`;
  const weeklyDiffColor = onTargetWeek ? C.accent : weeklyDiff > 0 ? C.warn : C.danger;

  const clientName = (id) => clients.find((c) => c.id === id)?.name || "—";
  const otherNoteRequired = form.category === "other";
  const missingOtherNote = otherNoteRequired && !form.note.trim();

  async function addEntry() {
    if (!form.clientId || !form.hours || Number(form.hours) <= 0 || missingOtherNote) return;
    const d = fromKey(form.date);
    const wk = `${user.id}|${toKey(startOfWeek(d))}`;
    if (lockedWeeks.includes(wk)) return;
    await supabase.from("time_entries").insert({
      employee_id: user.id, client_id: form.clientId, category: form.category,
      entry_date: form.date, hours: Number(form.hours), note: form.note.trim() || null,
    });
    setForm((f) => ({ ...f, hours: "", note: "" }));
    refetchEntries();
  }
  async function removeEntry(id) {
    await supabase.from("time_entries").delete().eq("id", id);
    refetchEntries();
  }
  async function saveEdit(id) {
    const val = Number(editHours);
    if (val > 0) {
      await supabase.from("time_entries").update({ hours: val, note: editNote.trim() || null }).eq("id", id);
      refetchEntries();
    }
    setEditingId(null);
  }
  async function lockWeek() {
    await supabase.from("locked_weeks").insert({ employee_id: user.id, week_start: toKey(weekStart) });
    refetchLockedWeeks();
  }

  return (
    <div>
      <PageHeader title="My Timesheet" sub="Log the hours you worked for each client, by day." />

      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="This week" value={weekTotal.toFixed(2) + "h"} accent={C.accent} />
        <StatCard label={`vs. ${weeklyTarget}h/week target`} value={weeklyDiffLabel} accent={weeklyDiffColor} />
        <StatCard label={fmtMonthYear(weekStart)} value={monthTotal.toFixed(2) + "h"} />
        <StatCard label="Entries this week" value={weekEntries.length} />
      </div>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
          <WeekNav weekStart={weekStart} setWeekStart={setWeekStart} />
          {!isLocked ? (
            <Btn variant="primary" size="sm" icon={Lock} onClick={lockWeek}>Submit week</Btn>
          ) : (
            <Badge tone="warn">Submitted &amp; locked</Badge>
          )}
        </div>
        {isLocked && (
          <div style={{
            display: "flex", alignItems: "center", gap: 8, background: C.warnSoft, color: C.warn,
            padding: "8px 12px", borderRadius: 7, fontFamily: sans, fontSize: 12.5, marginBottom: 14,
          }}>
            <Lock size={13} /> This week is submitted and locked. Ask your manager to reopen it if you need to make changes.
          </div>
        )}

        {!isLocked && (
          <div style={{
            display: "flex", gap: 10, alignItems: "flex-end", marginBottom: 18, flexWrap: "wrap",
            borderBottom: `1px solid ${C.border}`, paddingBottom: 18,
          }}>
            <Field label="Date">
              <input type="date" style={{ ...inputStyle, width: 150 }} value={form.date}
                onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))} />
            </Field>
            <Field label="Client">
              <select style={{ ...inputStyle, width: 190 }} value={form.clientId}
                onChange={(e) => setForm((f) => ({ ...f, clientId: e.target.value }))}>
                {clients.filter((c) => c.active).map((c) => (
                  <option key={c.id} value={c.id}>{c.name}</option>
                ))}
              </select>
            </Field>
            <Field label="Category">
              <select style={{ ...inputStyle, width: 140 }} value={form.category}
                onChange={(e) => setForm((f) => ({ ...f, category: e.target.value }))}>
                {CATEGORIES.map((cat) => <option key={cat} value={cat}>{CATEGORY_LABELS[cat]}</option>)}
              </select>
            </Field>
            <Field label="Hours">
              <input type="number" min="0.25" step="0.25" placeholder="e.g. 3.5" style={{ ...inputStyle, width: 90 }}
                value={form.hours} onChange={(e) => setForm((f) => ({ ...f, hours: e.target.value }))} />
            </Field>
            <Field label={otherNoteRequired ? "Note (required for \"Other\")" : "Note (optional)"}>
              <input type="text" placeholder="e.g. VAT return review" style={{ ...inputStyle, width: 200, ...(missingOtherNote ? { border: `1px solid ${C.danger}` } : {}) }}
                value={form.note} onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))} />
            </Field>
            <Btn icon={Plus} onClick={addEntry} disabled={missingOtherNote}>Add entry</Btn>
          </div>
        )}

        {days.map((day) => {
          const dayEntries = weekEntries.filter((e) => e.date === toKey(day));
          const dayTotal = dayEntries.reduce((s, e) => s + e.hours, 0);
          return (
            <div key={toKey(day)} style={{ display: "flex", gap: 16, padding: "10px 0", borderBottom: `1px solid ${C.border}` }}>
              <div style={{ width: 90, flexShrink: 0 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: isToday(day) ? C.accent : C.ink }}>
                  {fmtDow(day)} {fmtShort(day)}
                </div>
                {isWeekend(day) && <div style={{ fontFamily: sans, fontSize: 10.5, color: C.inkFaint }}>Weekend</div>}
              </div>
              <div style={{ flex: 1 }}>
                {dayEntries.length === 0 ? (
                  <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkFaint }}>No hours logged</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {dayEntries.map((e) => (
                      <div key={e.id}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <span style={{ fontFamily: sans, fontSize: 13, color: C.ink, minWidth: 170 }}>{clientName(e.clientId)}</span>
                          <Badge tone="neutral">{CATEGORY_LABELS[e.category] || "Other"}</Badge>
                          {editingId === e.id ? (
                            <>
                              <input type="number" min="0.25" step="0.25" autoFocus value={editHours}
                                onChange={(ev) => setEditHours(ev.target.value)}
                                style={{ ...inputStyle, width: 74, padding: "3px 7px", fontFamily: mono }} />
                              <input type="text" placeholder="Note (optional)" value={editNote}
                                onChange={(ev) => setEditNote(ev.target.value)}
                                style={{ ...inputStyle, width: 160, padding: "3px 7px" }} />
                              <button onClick={() => saveEdit(e.id)} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                              <button onClick={() => setEditingId(null)} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                            </>
                          ) : (
                            <>
                              <span style={{ fontFamily: mono, fontSize: 13, color: C.ink, fontWeight: 600 }}>{e.hours.toFixed(2)}h</span>
                              {!isLocked && (
                                <>
                                  <button onClick={() => { setEditingId(e.id); setEditHours(String(e.hours)); setEditNote(e.note || ""); }} style={iconBtnStyle}>
                                    <Pencil size={13} color={C.inkMuted} />
                                  </button>
                                  <button onClick={() => removeEntry(e.id)} style={iconBtnStyle}>
                                    <Trash2 size={13} color={C.inkMuted} />
                                  </button>
                                </>
                              )}
                            </>
                          )}
                        </div>
                        {editingId !== e.id && e.note && (
                          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint, marginTop: 2, marginLeft: 0 }}>{e.note}</div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div style={{ width: 60, textAlign: "right", fontFamily: mono, fontSize: 13, fontWeight: 600, color: dayTotal > 0 ? C.ink : C.inkFaint }}>
                {dayTotal > 0 ? dayTotal.toFixed(2) + "h" : "—"}
              </div>
            </div>
          );
        })}
        <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 12, fontFamily: sans, fontSize: 13, fontWeight: 700, color: C.ink }}>
          Week total:&nbsp;<span style={{ fontFamily: mono, color: C.accent }}>{weekTotal.toFixed(2)}h</span>
        </div>
      </Panel>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Time Off — request (employee) and history                              */
/* ---------------------------------------------------------------------- */

function TimeOff({ user, employees, leaveRequests, refetchLeaveRequests }) {
  const [form, setForm] = useState({ startDate: toKey(TODAY), endDate: toKey(TODAY), type: "annual", note: "" });

  const mine = leaveRequests
    .filter((r) => r.employeeId === user.id)
    .sort((a, b) => b.startDate.localeCompare(a.startDate));

  const requestYear = form.startDate ? fromKey(form.startDate).getFullYear() : TODAY.getFullYear();
  const holidayMap = holidayMapForYears([requestYear - 1, requestYear, requestYear + 1]);
  const usedThisYear = vacationDaysUsed(user.id, requestYear, leaveRequests, holidayMap);
  const annualLeaveDays = user.annualLeaveDays || ANNUAL_LEAVE_DAYS;
  const remaining = leaveBalanceThroughYear(user.id, annualLeaveDays, requestYear, leaveRequests);

  const validRange = form.startDate && form.endDate && form.endDate >= form.startDate;
  const requestedDays = validRange ? countWorkingDays(form.startDate, form.endDate, holidayMap) : 0;
  const overBalance = form.type === "annual" && requestedDays > remaining;
  const noteRequired = LEAVE_TYPES_REQUIRING_NOTE.has(form.type);
  const missingRequiredNote = noteRequired && !form.note.trim();
  const usageByType = daysUsedByType(user.id, requestYear, leaveRequests, holidayMap);
  const carryoverIn = leaveBalanceThroughYear(user.id, annualLeaveDays, requestYear - 1, leaveRequests);

  async function submit() {
    if (!validRange || overBalance || missingRequiredNote) return;
    await supabase.from("leave_requests").insert({
      employee_id: user.id, start_date: form.startDate, end_date: form.endDate,
      type: form.type, note: form.note.trim() || null, status: "pending",
    });
    setForm((f) => ({ ...f, note: "" }));
    refetchLeaveRequests();
  }
  async function cancelRequest(id) {
    await supabase.from("leave_requests").delete().eq("id", id);
    refetchLeaveRequests();
  }
  const decidedByName = (id) => employees.find((e) => e.id === id)?.name || "—";

  return (
    <div>
      <PageHeader title="Time Off" sub="Request leave and track approval status." />
      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label={`Annual leave used in ${requestYear}`} value={`${usedThisYear} / ${annualLeaveDays}`} sub="working days" />
        <StatCard label={`Carried over from ${requestYear - 1}`} value={`${carryoverIn >= 0 ? "+" : ""}${carryoverIn}d`} accent={carryoverIn < 0 ? C.danger : undefined} />
        <StatCard label={`Balance through ${requestYear}`} value={`${remaining >= 0 ? "+" : ""}${remaining}d`} accent={remaining < 0 ? C.danger : remaining === 0 ? C.warn : C.accent} sub="carries over year to year" />
      </div>
      <div style={{ display: "flex", gap: 18, flexWrap: "wrap" }}>
        <Panel style={{ flex: "0 0 320px" }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 14 }}>New request</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Start date">
              <input type="date" style={inputStyle} value={form.startDate} onChange={(e) => setForm((f) => ({ ...f, startDate: e.target.value }))} />
            </Field>
            <Field label="End date">
              <input type="date" style={inputStyle} value={form.endDate} onChange={(e) => setForm((f) => ({ ...f, endDate: e.target.value }))} />
            </Field>
            <Field label="Type">
              <select style={inputStyle} value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
                {Object.entries(LEAVE_TYPES).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
            </Field>
            <Field label={noteRequired ? "Note (required for this leave type)" : "Note (optional)"}>
              <input style={{ ...inputStyle, ...(missingRequiredNote ? { border: `1px solid ${C.danger}` } : {}) }} value={form.note} onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))} placeholder="e.g. Family trip" />
            </Field>
            {noteRequired && (
              <div style={{ fontFamily: sans, fontSize: 12, color: missingRequiredNote ? C.danger : C.inkMuted }}>
                This leave type requires a note explaining the reason.
              </div>
            )}
            {validRange && form.type === "annual" && (
              <div style={{ fontFamily: sans, fontSize: 12, color: overBalance ? C.danger : C.inkMuted }}>
                Uses {requestedDays} working day{requestedDays === 1 ? "" : "s"} of your {remaining}d balance
                {overBalance ? " — exceeds your balance." : "."}
              </div>
            )}
            <Btn icon={Plus} onClick={submit} disabled={!validRange || overBalance || missingRequiredNote}>Submit request</Btn>
          </div>
        </Panel>
        <Panel style={{ flex: "0 0 220px" }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Used in {requestYear}</div>
          <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, marginBottom: 10 }}>By leave type, approved only</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Object.entries(LEAVE_TYPES).map(([key, label]) => (
              <div key={key} style={{ display: "flex", justifyContent: "space-between", fontFamily: sans, fontSize: 12.5 }}>
                <span style={{ color: C.inkMuted }}>{label}</span>
                <span style={{ fontFamily: mono, fontWeight: 600, color: C.ink }}>{usageByType[key] || 0}d</span>
              </div>
            ))}
          </div>
        </Panel>
        <Panel style={{ flex: 1, minWidth: 320 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Your requests</div>
          {mine.length === 0 ? <EmptyState icon={CalendarDays} title="No time off requested yet" /> : (
            <TableShell headers={["Dates", "Type", "Note", "Status", ""]}>
              {mine.map((r) => (
                <tr key={r.id}>
                  <Td>{fmtShort(fromKey(r.startDate))} – {fmtShort(fromKey(r.endDate))}</Td>
                  <Td>{leaveTypeLabel(r.type)}</Td>
                  <Td>{r.note || "—"}</Td>
                  <Td>
                    <Badge tone={leaveStatusTone(r.status)}>
                      {r.status === "approved" ? `Approved by ${decidedByName(r.decidedBy)}` : r.status === "rejected" ? `Rejected by ${decidedByName(r.decidedBy)}` : "Pending"}
                    </Badge>
                  </Td>
                  <Td right>
                    {r.status === "pending" && (
                      <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => cancelRequest(r.id)}>Cancel</Btn>
                    )}
                  </Td>
                </tr>
              ))}
            </TableShell>
          )}
        </Panel>
      </div>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Leave Approvals                                                  */
/* ---------------------------------------------------------------------- */

function LeaveApprovals({ user, employees, leaveRequests, refetchLeaveRequests }) {
  const [statusFilter, setStatusFilter] = useState("pending");
  const [year, setYear] = useState(TODAY.getFullYear());
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";

  const holidayMap = holidayMapForYears([year - 1, year, year + 1]);
  const balances = employees.filter((e) => e.active).map((e) => {
    const entitlement = e.annualLeaveDays || ANNUAL_LEAVE_DAYS;
    const used = vacationDaysUsed(e.id, year, leaveRequests, holidayMap);
    const carryover = leaveBalanceThroughYear(e.id, entitlement, year - 1, leaveRequests);
    const balance = carryover + entitlement - used;
    const byType = daysUsedByType(e.id, year, leaveRequests, holidayMap);
    return { ...e, entitlement, used, carryover, balance, byType };
  }).sort((a, b) => a.balance - b.balance);

  const rows = leaveRequests
    .filter((r) => statusFilter === "all" || r.status === statusFilter)
    .sort((a, b) => b.startDate.localeCompare(a.startDate));

  async function decide(id, status) {
    await supabase.from("leave_requests").update({ status, decided_by: user.id, decided_at: toKey(TODAY) }).eq("id", id);
    refetchLeaveRequests();
  }

  return (
    <div>
      <PageHeader title="Leave Approvals" sub="Review and decide on time off requests from your team." />

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 4 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Annual leave balances</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={() => setYear((y) => y - 1)} style={navBtnStyle}><ChevronLeft size={16} /></button>
            <div style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, minWidth: 50, textAlign: "center" }}>{year}</div>
            <button onClick={() => setYear((y) => y + 1)} style={navBtnStyle}><ChevronRight size={16} /></button>
          </div>
        </div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
          Balance carries over from year to year — a negative number means more days were taken than earned. Public holidays and weekends don't count against it.
        </div>
        <TableShell headers={["Employee", `Carried over from ${year - 1}`, "Entitlement", `Used in ${year}`, `Balance through ${year}`, ""]}>
          {balances.map((e) => {
            const pct = Math.min(100, (e.used / e.entitlement) * 100);
            const tone = e.balance < 0 ? "danger" : e.balance <= 5 ? "warn" : "accent";
            return (
              <tr key={e.id}>
                <Td>{e.name}</Td>
                <Td mono style={{ color: e.carryover < 0 ? C.danger : undefined }}>{e.carryover >= 0 ? "+" : ""}{e.carryover}d</Td>
                <Td mono>{e.entitlement}d</Td>
                <Td mono>{e.used}d</Td>
                <Td mono style={{ color: e.balance < 0 ? C.danger : undefined, fontWeight: 700 }}>{e.balance >= 0 ? "+" : ""}{e.balance}d</Td>
                <Td>
                  <div style={{ width: 120, height: 6, borderRadius: 999, background: C.border, overflow: "hidden" }}>
                    <div style={{ width: `${pct}%`, height: "100%", background: tone === "danger" ? C.danger : tone === "warn" ? C.warn : C.accent }} />
                  </div>
                </Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Leave usage by type — {year}</div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Approved working days, per employee, per category.</div>
        <TableShell headers={["Employee", ...Object.values(LEAVE_TYPES)]}>
          {balances.map((e) => (
            <tr key={e.id}>
              <Td>{e.name}</Td>
              {Object.keys(LEAVE_TYPES).map((key) => (
                <Td key={key} mono>{e.byType[key] || 0}d</Td>
              ))}
            </tr>
          ))}
        </TableShell>
      </Panel>

      <Panel>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {["pending", "approved", "rejected", "all"].map((s) => (
            <button key={s} onClick={() => setStatusFilter(s)}
              style={{
                padding: "6px 12px", borderRadius: 999, border: `1px solid ${C.border}`, cursor: "pointer",
                fontFamily: sans, fontSize: 12.5, fontWeight: 600, textTransform: "capitalize",
                background: statusFilter === s ? C.accent : C.surface, color: statusFilter === s ? "#fff" : C.inkMuted,
              }}>
              {s}
            </button>
          ))}
        </div>
        {rows.length === 0 ? <EmptyState icon={ClipboardCheck} title="No requests here" /> : (
          <TableShell headers={["Employee", "Dates", "Type", "Note", "Status", ""]}>
            {rows.map((r) => (
              <tr key={r.id}>
                <Td>{empName(r.employeeId)}</Td>
                <Td>{fmtShort(fromKey(r.startDate))} – {fmtShort(fromKey(r.endDate))}</Td>
                <Td>{leaveTypeLabel(r.type)}</Td>
                <Td>{r.note || "—"}</Td>
                <Td><Badge tone={leaveStatusTone(r.status)}>{r.status}</Badge></Td>
                <Td right>
                  {r.status === "pending" ? (
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <Btn size="sm" variant="secondary" icon={Check} onClick={() => decide(r.id, "approved")}>Approve</Btn>
                      <Btn size="sm" variant="danger" icon={X} onClick={() => decide(r.id, "rejected")}>Reject</Btn>
                    </div>
                  ) : (
                    <span style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint }}>by {empName(r.decidedBy)}</span>
                  )}
                </Td>
              </tr>
            ))}
          </TableShell>
        )}
      </Panel>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Company Calendar — public holidays + approved leave                    */
/* ---------------------------------------------------------------------- */

const WEEKDAY_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

function buildMonthGrid(monthAnchor) {
  const first = startOfMonth(monthAnchor);
  const gridStart = startOfWeek(first);
  const days = [];
  for (let i = 0; i < 42; i++) days.push(addDays(gridStart, i));
  return days;
}

function TeamCalendar({ employees, leaveRequests }) {
  const [monthAnchor, setMonthAnchor] = useState(startOfMonth(TODAY));
  const year = monthAnchor.getFullYear();
  const holidayMap = holidayMapForYears([year - 1, year, year + 1]);
  const days = buildMonthGrid(monthAnchor);
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const initials = (id) => (empName(id).match(/\b\w/g) || []).slice(0, 2).join("").toUpperCase();

  const approvedLeave = leaveRequests.filter((r) => r.status === "approved");
  function leaveOn(day) {
    const k = toKey(day);
    return approvedLeave.filter((r) => r.startDate <= k && k <= r.endDate);
  }

  return (
    <div>
      <PageHeader title="Calendar" sub="Public holidays and approved time off, at a glance."
        right={
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={() => setMonthAnchor(new Date(year, monthAnchor.getMonth() - 1, 1))} style={navBtnStyle}><ChevronLeft size={16} /></button>
            <div style={{ fontFamily: sans, fontSize: 13.5, fontWeight: 600, minWidth: 140, textAlign: "center" }}>{fmtMonthYear(monthAnchor)}</div>
            <button onClick={() => setMonthAnchor(new Date(year, monthAnchor.getMonth() + 1, 1))} style={navBtnStyle}><ChevronRight size={16} /></button>
            <Btn variant="ghost" size="sm" onClick={() => setMonthAnchor(startOfMonth(new Date()))}>This month</Btn>
          </div>
        } />
      <Panel>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6, marginBottom: 6 }}>
          {WEEKDAY_LABELS.map((d) => (
            <div key={d} style={{ fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.inkFaint, textAlign: "center", padding: "2px 0" }}>{d}</div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6 }}>
          {days.map((day) => {
            const inMonth = day.getMonth() === monthAnchor.getMonth();
            const holiday = holidayMap.get(toKey(day));
            const dayLeave = leaveOn(day);
            return (
              <div key={toKey(day)} style={{
                minHeight: 78, borderRadius: 8, padding: 6,
                border: `1px solid ${C.border}`,
                background: holiday ? C.warnSoft : isToday(day) ? C.accentSoft : C.surface,
                opacity: inMonth ? 1 : 0.4,
              }}>
                <div style={{ fontFamily: mono, fontSize: 11.5, fontWeight: 600, color: isToday(day) ? C.accentDark : C.ink }}>{day.getDate()}</div>
                {holiday && <div style={{ fontFamily: sans, fontSize: 10, color: C.warn, marginTop: 2, lineHeight: 1.3 }}>{holiday}</div>}
                {dayLeave.length > 0 && (
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 3, marginTop: 4 }}>
                    {dayLeave.map((r) => (
                      <span key={r.id} title={`${empName(r.employeeId)} — ${leaveTypeLabel(r.type)}`} style={{
                        fontFamily: sans, fontSize: 9.5, fontWeight: 700, color: C.accentDark, background: C.accentSoft,
                        borderRadius: 999, padding: "1px 5px",
                      }}>
                        {initials(r.employeeId)}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <div style={{ display: "flex", gap: 16, marginTop: 14, fontFamily: sans, fontSize: 11.5, color: C.inkMuted }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ width: 12, height: 12, borderRadius: 3, background: C.warnSoft, border: `1px solid ${C.border}` }} />
            Public holiday
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ fontFamily: sans, fontSize: 9.5, fontWeight: 700, color: C.accentDark, background: C.accentSoft, borderRadius: 999, padding: "1px 5px" }}>AB</span>
            Approved time off (initials)
          </div>
        </div>
      </Panel>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Admin: Dashboard                                                        */
/* ---------------------------------------------------------------------- */

function AdminDashboard({ employees, clients, entries }) {
  const [period, setPeriod] = useState("month"); // "week" | "month" | "quarter" | "year"
  const [anchor, setAnchor] = useState(TODAY);
  const DEFAULT_DASHBOARD_CONFIG = {
    statCards: true, allocationChart: true, revenueCostChart: true,
    byClientTable: true, byEmployeeTable: true,
    monthlyEmployeeTable: true, weeklyEmployeeTable: true, monthlyClientTable: true, companyMonthlyTable: true,
  };
  const [configRaw, setConfig] = useAppSetting("dashboard_config", DEFAULT_DASHBOARD_CONFIG);
  const config = configRaw || DEFAULT_DASHBOARD_CONFIG;
  const [customizing, setCustomizing] = useState(false);

  const [rangeStart, rangeEnd] = periodRange(period, anchor);

  const rangeEntries = entries.filter((e) => {
    const d = fromKey(e.date);
    return d >= rangeStart && d <= rangeEnd;
  });

  const activeEmployees = new Set(rangeEntries.map((e) => e.employeeId)).size;

  // Each employee's own hours and labor cost for the period, computed
  // once and reused below both to attribute cost to clients and to show
  // per-employee profitability.
  const employeeCosts = employees.filter((e) => e.active && (!earliestHistoryDate(e.costHistory) || earliestHistoryDate(e.costHistory) <= toKey(rangeEnd))).map((emp) => ({
    id: emp.id,
    hours: rangeEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0),
    cost: periodEmployeeCost(emp, rangeStart, rangeEnd),
  }));

  const byClient = clients.filter((c) => c.active && (!earliestHistoryDate(c.feeHistory) || earliestHistoryDate(c.feeHistory) <= toKey(rangeEnd))).map((c) => {
    const clientEntries = rangeEntries.filter((e) => e.clientId === c.id);
    const hrs = clientEntries.reduce((s, e) => s + e.hours, 0);
    const { revenue, allocation } = periodClientMetrics(c, clientEntries, rangeStart, rangeEnd);
    // Attribute each employee's labor cost to this client proportionally
    // to the hours they spent on it — the mirror image of how revenue is
    // attributed to employees below.
    const cost = employeeCosts.reduce((sum, ec) => {
      if (ec.hours <= 0) return sum;
      const empHrsForClient = clientEntries.filter((e) => e.employeeId === ec.id).reduce((s, e) => s + e.hours, 0);
      return sum + ec.cost * (empHrsForClient / ec.hours);
    }, 0);
    const profit = revenue - cost;
    const margin = revenue > 0 ? (profit / revenue) * 100 : null;
    return { ...c, hours: hrs, allocation, revenue, effectiveRate: hrs > 0 ? revenue / hrs : null, cost, profit, margin };
  }).sort((a, b) => b.hours - a.hours);

  const totalRevenue = byClient.reduce((s, c) => s + c.revenue, 0);
  const totalHours = rangeEntries.reduce((s, e) => s + e.hours, 0);
  const avgEffectiveRate = totalHours > 0 ? totalRevenue / totalHours : 0;

  // Attribute each client's revenue to employees proportionally to the
  // hours they logged for that client — a simple productivity signal,
  // not a payroll or commission calculation.
  const byEmployee = employees.filter((e) => e.active && (!earliestHistoryDate(e.costHistory) || earliestHistoryDate(e.costHistory) <= toKey(rangeEnd))).map((emp) => {
    const hrs = rangeEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
    const clientsTouched = new Set(rangeEntries.filter((e) => e.employeeId === emp.id).map((e) => e.clientId)).size;
    const revenue = byClient.reduce((sum, c) => {
      const clientTotalHrs = c.hours;
      if (clientTotalHrs <= 0) return sum;
      const empHrsForClient = rangeEntries.filter((e) => e.employeeId === emp.id && e.clientId === c.id).reduce((s, e) => s + e.hours, 0);
      return sum + c.revenue * (empHrsForClient / clientTotalHrs);
    }, 0);
    const cost = employeeCosts.find((ec) => ec.id === emp.id)?.cost || 0;
    const profit = revenue - cost;
    const margin = revenue > 0 ? (profit / revenue) * 100 : null;
    return { ...emp, hours: hrs, clientsTouched, revenue, effectiveRate: hrs > 0 ? revenue / hrs : null, cost, profit, margin };
  }).sort((a, b) => b.hours - a.hours);

  const totalCost = byEmployee.reduce((s, e) => s + e.cost, 0);
  const totalProfit = totalRevenue - totalCost;
  const overallMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : null;

  // Clients/employees that first started within THIS period specifically
  // (not just "existing by now") — driven by the same "Effective from" /
  // cost-history date you set, so it reflects what you actually told the
  // app, not when the database row happened to be created.
  const newClientsThisPeriod = clients.filter((c) => {
    const start = earliestHistoryDate(c.feeHistory);
    if (!start) return false;
    const d = fromKey(start);
    return d >= rangeStart && d <= rangeEnd;
  });
  const newEmployeesThisPeriod = employees.filter((e) => {
    const start = earliestHistoryDate(e.costHistory);
    if (!start) return false;
    const d = fromKey(start);
    return d >= rangeStart && d <= rangeEnd;
  });
  const newHiresCost = newEmployeesThisPeriod.reduce((sum, e) => {
    const ec = employeeCosts.find((x) => x.id === e.id);
    return sum + (ec ? ec.cost : 0);
  }, 0);

  // Rolling 6-month window (oldest to newest, ending on the anchor month)
  // for the per-employee, per-client, and company-wide monthly trends.
  // Each month uses whichever fee/cost terms were actually in effect.
  const trendMonths = Array.from({ length: 6 }, (_, i) => {
    const d = new Date(anchor.getFullYear(), anchor.getMonth() - (5 - i), 1);
    return { label: d.toLocaleDateString("en-GB", { month: "short", year: "2-digit" }), start: startOfMonth(d), end: endOfMonth(d) };
  });
  const employeeMonthlyHours = employees.filter((e) => e.active).map((emp) => ({
    ...emp,
    months: trendMonths.map((m) => entries.filter((e) => e.employeeId === emp.id && fromKey(e.date) >= m.start && fromKey(e.date) <= m.end).reduce((s, e) => s + e.hours, 0)),
  }));

  // Rolling 8-week window (oldest to newest, ending on the anchor's week)
  // for the per-employee weekly trend table.
  const trendWeeks = Array.from({ length: 8 }, (_, i) => {
    const start = addDays(startOfWeek(anchor), (i - 7) * 7);
    return { label: fmtShort(start), start, end: addDays(start, 6) };
  });
  const employeeWeeklyHours = employees.filter((e) => e.active).map((emp) => ({
    ...emp,
    weeks: trendWeeks.map((w) => entries.filter((e) => e.employeeId === emp.id && fromKey(e.date) >= w.start && fromKey(e.date) <= w.end).reduce((s, e) => s + e.hours, 0)),
  }));

  const clientMonthly = clients.filter((c) => c.active).map((c) => {
    const months = trendMonths.map((m) => {
      const monthEntries = entries.filter((e) => e.clientId === c.id && fromKey(e.date) >= m.start && fromKey(e.date) <= m.end);
      const hrs = monthEntries.reduce((s, e) => s + e.hours, 0);
      return { hours: hrs, revenue: periodClientMetrics(c, monthEntries, m.start, m.end).revenue };
    });
    return { ...c, months, totalRevenue: months.reduce((s, m) => s + m.revenue, 0) };
  });

  // Company-wide monthly overview: new clients added, revenue, cost, profit.
  const companyMonthly = trendMonths.map((m) => {
    const newClients = clients.filter((c) => c.createdAt && fromKey(c.createdAt) >= m.start && fromKey(c.createdAt) <= m.end).length;
    const monthEntries = entries.filter((e) => fromKey(e.date) >= m.start && fromKey(e.date) <= m.end);
    const revenue = clients.filter((c) => c.active).reduce((sum, c) => sum + periodClientMetrics(c, monthEntries.filter((e) => e.clientId === c.id), m.start, m.end).revenue, 0);
    const cost = employees.filter((e) => e.active).reduce((sum, e) => sum + periodEmployeeCost(e, m.start, m.end), 0);
    return { label: m.label, newClients, revenue, cost, profit: revenue - cost };
  });

  const chartData = byClient.filter((c) => c.hours > 0 || c.allocation).map((c) => ({
    name: c.name.length > 14 ? c.name.slice(0, 13) + "…" : c.name,
    Actual: Number(c.hours.toFixed(2)),
    Allocated: c.allocation ? Number(c.allocation.toFixed(2)) : 0,
  }));

  function shiftAnchor(dir) {
    setAnchor((a) => shiftPeriodAnchor(period, a, dir));
  }

  function exportReport() {
    exportToExcel(`dashboard-report-${toKey(TODAY)}.xlsx`, [
      { name: "By client", rows: byClient.map((c) => ({
        Client: c.name, Hours: Number(c.hours.toFixed(2)), "Allocation (h)": c.allocation ? Number(c.allocation.toFixed(1)) : "",
        "Revenue (€)": Number(c.revenue.toFixed(2)), "Cost (€)": Number(c.cost.toFixed(2)), "Profit (€)": Number(c.profit.toFixed(2)),
        "Margin (%)": c.margin !== null ? Number(c.margin.toFixed(1)) : "",
      })) },
      { name: "By employee", rows: byEmployee.map((e) => ({
        Employee: e.name, Hours: Number(e.hours.toFixed(2)), Clients: e.clientsTouched,
        "Revenue (€)": Number(e.revenue.toFixed(2)), "Cost (€)": Number(e.cost.toFixed(2)), "Profit (€)": Number(e.profit.toFixed(2)),
        "Margin (%)": e.margin !== null ? Number(e.margin.toFixed(1)) : "",
      })) },
      { name: "Company by month", rows: companyMonthly.map((m) => ({
        Month: m.label, "New clients": m.newClients, "Revenue (€)": Number(m.revenue.toFixed(2)),
        "Cost (€)": Number(m.cost.toFixed(2)), "Profit (€)": Number(m.profit.toFixed(2)),
      })) },
    ]);
  }

  const CONFIG_LABELS = {
    statCards: "Summary cards", allocationChart: "Actual vs. allocated chart", revenueCostChart: "Revenue vs. cost chart",
    byClientTable: "Hours by client table", byEmployeeTable: "Hours by employee table",
    monthlyEmployeeTable: "Hours per employee, by month", weeklyEmployeeTable: "Hours per employee, by week", monthlyClientTable: "Client analysis, by month",
    companyMonthlyTable: "Company overview, by month",
  };

  return (
    <div>
      <PageHeader
        title="Dashboard"
        sub="Hours across the company, by employee and by client."
        right={
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
              {PERIOD_OPTIONS.map((p) => (
                <button key={p.key} onClick={() => setPeriod(p.key)}
                  style={{
                    padding: "7px 11px", fontSize: 12.5, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer",
                    background: period === p.key ? C.accent : C.surface, color: period === p.key ? "#fff" : C.inkMuted,
                  }}>
                  {p.label}
                </button>
              ))}
            </div>
            <button onClick={() => shiftAnchor(-1)} style={navBtnStyle}><ChevronLeft size={16} /></button>
            <div style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, minWidth: 150, textAlign: "center" }}>
              {periodLabel(period, rangeStart, rangeEnd)}
            </div>
            <button onClick={() => shiftAnchor(1)} style={navBtnStyle}><ChevronRight size={16} /></button>
            <Btn variant="secondary" size="sm" onClick={() => setCustomizing(true)}>Customize</Btn>
            <Btn variant="secondary" size="sm" onClick={exportReport}>Export to Excel</Btn>
          </div>
        }
      />

      {config.statCards && (
        <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
          <StatCard label="Active employees" value={activeEmployees} sub={`of ${employees.filter(e=>e.role==='employee' && (!earliestHistoryDate(e.costHistory) || earliestHistoryDate(e.costHistory) <= toKey(rangeEnd))).length} employees`} />
          <StatCard label="Clients with hours" value={byClient.filter((c) => c.hours > 0).length} sub={`of ${byClient.length} active clients`} />
          <StatCard label="New clients this period" value={newClientsThisPeriod.length} sub={newClientsThisPeriod.length ? newClientsThisPeriod.map(c=>c.name).join(", ") : "None started this period"} />
          <StatCard label="New employees this period" value={newEmployeesThisPeriod.length} sub={newEmployeesThisPeriod.length ? fmtEur(newHiresCost) + " added to labor cost" : "None started this period"} />
          <StatCard label="Revenue this period" value={fmtEur(totalRevenue)} accent={C.accent} sub="Fixed fees + extra fees" />
          <StatCard label="Labor cost this period" value={fmtEur(totalCost)} sub="Gross + benefits, prorated" />
          <StatCard label="Profit this period" value={fmtEur(totalProfit)} accent={totalProfit < 0 ? C.danger : C.accent} sub="Revenue − cost" />
          <StatCard label="Overall margin" value={overallMargin !== null ? overallMargin.toFixed(0) + "%" : "—"} accent={overallMargin !== null && overallMargin < 0 ? C.danger : C.accent} sub="(Revenue − cost) ÷ revenue" />
          <StatCard label="Avg. effective rate" value={fmtEur(avgEffectiveRate) + "/h"} sub="Revenue ÷ hours worked" />
        </div>
      )}

      {config.allocationChart && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 14 }}>
            Actual vs. allocated hours per client
          </div>
          {chartData.length === 0 ? (
            <EmptyState icon={Building2} title="No hours in this period" />
          ) : (
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <BarChart data={chartData} margin={{ left: -10, right: 10 }}>
                  <CartesianGrid stroke={C.border} vertical={false} />
                  <XAxis dataKey="name" tick={{ fontFamily: sans, fontSize: 11, fill: C.inkMuted }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis tick={{ fontFamily: mono, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: sans, fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} />
                  <Bar dataKey="Actual" fill={C.accent} radius={[4, 4, 0, 0]} maxBarSize={26} />
                  <Bar dataKey="Allocated" fill={C.border} radius={[4, 4, 0, 0]} maxBarSize={26} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </Panel>
      )}

      {config.revenueCostChart && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Revenue vs. cost, by month</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 14 }}>Last 6 months, ending {trendMonths[5].label}.</div>
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <LineChart data={companyMonthly} margin={{ left: -10, right: 10 }}>
                <CartesianGrid stroke={C.border} vertical={false} />
                <XAxis dataKey="label" tick={{ fontFamily: sans, fontSize: 11, fill: C.inkMuted }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontFamily: mono, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ fontFamily: sans, fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} formatter={(v) => fmtEur(v)} />
                <Legend wrapperStyle={{ fontFamily: sans, fontSize: 12 }} />
                <Line type="monotone" dataKey="revenue" name="Revenue" stroke={C.accent} strokeWidth={2.5} dot={{ r: 3 }} />
                <Line type="monotone" dataKey="cost" name="Cost" stroke={C.danger} strokeWidth={2.5} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </Panel>
      )}

      {config.byClientTable && (
        <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Hours by client</div>
            <TableShell headers={["Client", "Hours", "Allocation", "Status", "Revenue", "Cost", "Profit", "Margin"]}>
              {byClient.map((c) => {
                const pct = c.allocation ? (c.hours / c.allocation) * 100 : null;
                let tone = "neutral", label = "No allocation set";
                if (pct !== null) {
                  if (pct > 100) { tone = "danger"; label = `${pct.toFixed(0)}% — over`; }
                  else if (pct >= 85) { tone = "warn"; label = `${pct.toFixed(0)}% — near limit`; }
                  else { tone = "accent"; label = `${pct.toFixed(0)}% used`; }
                }
                return (
                  <tr key={c.id}>
                    <Td>{c.name}</Td>
                    <Td mono>{c.hours.toFixed(2)}h</Td>
                    <Td mono>{c.allocation ? c.allocation.toFixed(1) + "h" : "—"}</Td>
                    <Td><Badge tone={tone}>{label}</Badge></Td>
                    <Td mono>{fmtEur(c.revenue)}</Td>
                    <Td mono>{fmtEur(c.cost)}</Td>
                    <Td mono style={{ color: c.profit < 0 ? C.danger : undefined }}>{fmtEur(c.profit)}</Td>
                    <Td>{c.margin !== null ? <Badge tone={c.margin < 0 ? "danger" : c.margin < 20 ? "warn" : "accent"}>{c.margin.toFixed(0)}%</Badge> : "—"}</Td>
                  </tr>
                );
              })}
            </TableShell>
        </Panel>
      )}

      {config.byEmployeeTable && (
        <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Hours by employee</div>
            <TableShell headers={["Employee", "Hours", "Clients", "Revenue", "Cost", "Profit", "Margin"]}>
              {byEmployee.map((e) => (
                <tr key={e.id}>
                  <Td>{e.name}</Td>
                  <Td mono>{e.hours.toFixed(2)}h</Td>
                  <Td mono>{e.clientsTouched}</Td>
                  <Td mono>{fmtEur(e.revenue)}</Td>
                  <Td mono>{fmtEur(e.cost)}</Td>
                  <Td mono style={{ color: e.profit < 0 ? C.danger : undefined }}>{fmtEur(e.profit)}</Td>
                  <Td>{e.margin !== null ? <Badge tone={e.margin < 0 ? "danger" : e.margin < 20 ? "warn" : "accent"}>{e.margin.toFixed(0)}%</Badge> : "—"}</Td>
                </tr>
              ))}
            </TableShell>
        </Panel>
      )}

      {config.companyMonthlyTable && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Company overview, by month</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>New clients, revenue, cost, and profit — last 6 months.</div>
          <TableShell headers={["Month", "New clients", "Revenue", "Cost", "Profit"]}>
            {companyMonthly.map((m, i) => (
              <tr key={i}>
                <Td>{m.label}</Td>
                <Td mono>{m.newClients}</Td>
                <Td mono>{fmtEur(m.revenue)}</Td>
                <Td mono>{fmtEur(m.cost)}</Td>
                <Td mono style={{ color: m.profit < 0 ? C.danger : undefined }}>{fmtEur(m.profit)}</Td>
              </tr>
            ))}
          </TableShell>
        </Panel>
      )}

      <div style={{ display: "flex", gap: 18, flexWrap: "wrap", marginTop: 18 }}>
        {config.monthlyEmployeeTable && (
          <Panel style={{ flex: 1, minWidth: 380 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Hours per employee, by month</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Last 6 months, ending {trendMonths[5].label}.</div>
            <TableShell headers={["Employee", ...trendMonths.map((m) => m.label), "Total"]}>
              {employeeMonthlyHours.map((e) => (
                <tr key={e.id}>
                  <Td>{e.name}</Td>
                  {e.months.map((h, i) => <Td key={i} mono>{h > 0 ? h.toFixed(1) : "—"}</Td>)}
                  <Td mono>{e.months.reduce((s, h) => s + h, 0).toFixed(1)}</Td>
                </tr>
              ))}
            </TableShell>
          </Panel>
        )}

        {config.monthlyClientTable && (
          <Panel style={{ flex: 1, minWidth: 380 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Client analysis, by month</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Hours per month, plus total revenue over the window.</div>
            <TableShell headers={["Client", ...trendMonths.map((m) => m.label), "Revenue (6mo)"]}>
              {clientMonthly.map((c) => (
                <tr key={c.id}>
                  <Td>{c.name}</Td>
                  {c.months.map((m, i) => <Td key={i} mono>{m.hours > 0 ? m.hours.toFixed(1) : "—"}</Td>)}
                  <Td mono>{fmtEur(c.totalRevenue)}</Td>
                </tr>
              ))}
            </TableShell>
          </Panel>
        )}
      </div>

      {config.weeklyEmployeeTable && (
        <Panel style={{ marginTop: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Hours per employee, by week</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Last 8 weeks (starting date shown), ending the week of {trendWeeks[7].label}.</div>
          <TableShell headers={["Employee", ...trendWeeks.map((w) => w.label), "Total"]}>
            {employeeWeeklyHours.map((e) => (
              <tr key={e.id}>
                <Td>{e.name}</Td>
                {e.weeks.map((h, i) => <Td key={i} mono>{h > 0 ? h.toFixed(1) : "—"}</Td>)}
                <Td mono>{e.weeks.reduce((s, h) => s + h, 0).toFixed(1)}</Td>
              </tr>
            ))}
          </TableShell>
        </Panel>
      )}

      {customizing && (
        <Modal title="Customize dashboard" onClose={() => setCustomizing(false)} width={380}>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, marginBottom: 4 }}>
              Choose which sections to show. This applies for everyone using the dashboard.
            </div>
            {Object.keys(CONFIG_LABELS).map((key) => (
              <label key={key} style={{ display: "flex", alignItems: "center", gap: 10, fontFamily: sans, fontSize: 13.5, color: C.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={!!config[key]} onChange={(e) => setConfig({ ...config, [key]: e.target.checked })} />
                {CONFIG_LABELS[key]}
              </label>
            ))}
            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 10 }}>
              <Btn onClick={() => setCustomizing(false)}>Done</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}



/* ---------------------------------------------------------------------- */
/* Admin: All timesheets (view + edit any entry)                          */
/* ---------------------------------------------------------------------- */

function AdminTimesheets({ employees, clients, entries, refetchEntries, lockedWeeks, refetchLockedWeeks }) {
  const [weekStart, setWeekStart] = useState(startOfWeek(TODAY));
  const [empFilter, setEmpFilter] = useState("all");
  const [clientFilter, setClientFilter] = useState("all");
  const [editingId, setEditingId] = useState(null);
  const [editHours, setEditHours] = useState("");
  const [editNote, setEditNote] = useState("");

  const weekEnd = addDays(weekStart, 6);
  const clientName = (id) => clients.find((c) => c.id === id)?.name || "—";
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";

  const rows = entries
    .filter((e) => { const d = fromKey(e.date); return d >= weekStart && d <= weekEnd; })
    .filter((e) => empFilter === "all" || e.employeeId === empFilter)
    .filter((e) => clientFilter === "all" || e.clientId === clientFilter)
    .sort((a, b) => a.date.localeCompare(b.date) || empName(a.employeeId).localeCompare(empName(b.employeeId)));

  async function saveEdit(id) {
    const val = Number(editHours);
    if (val > 0) { await supabase.from("time_entries").update({ hours: val, note: editNote.trim() || null }).eq("id", id); refetchEntries(); }
    setEditingId(null);
  }
  async function removeEntry(id) { await supabase.from("time_entries").delete().eq("id", id); refetchEntries(); }
  function isRowLocked(e) { return lockedWeeks.includes(`${e.employeeId}|${toKey(startOfWeek(fromKey(e.date)))}`); }
  async function toggleRowLock(e) {
    const wk = toKey(startOfWeek(fromKey(e.date)));
    if (isRowLocked(e)) await supabase.from("locked_weeks").delete().eq("employee_id", e.employeeId).eq("week_start", wk);
    else await supabase.from("locked_weeks").insert({ employee_id: e.employeeId, week_start: wk });
    refetchLockedWeeks();
  }

  const total = rows.reduce((s, e) => s + e.hours, 0);

  const weekEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= weekStart && d <= weekEnd; });
  const compliance = employees.filter((e) => e.active).map((e) => {
    const target = e.weeklyHours || 40;
    const logged = weekEntries.filter((t) => t.employeeId === e.id).reduce((s, t) => s + t.hours, 0);
    const diff = logged - target;
    return { ...e, target, logged, diff };
  }).sort((a, b) => b.diff - a.diff);

  return (
    <div>
      <PageHeader title="All Timesheets" sub="Review and correct any employee's logged hours." />

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>
          Weekly hours vs. contract
        </div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
          {fmtShort(weekStart)} – {fmtShort(weekEnd)}. Each employee's contracted hours are set on the Employees page.
        </div>
        <TableShell headers={["Employee", "Contracted /wk", "Logged", "Difference", "Status"]}>
          {compliance.map((e) => {
            const onTrack = Math.abs(e.diff) < 0.01;
            const over = e.diff > 0;
            const tone = onTrack ? "accent" : over ? "warn" : "danger";
            const label = onTrack ? "On track" : over ? `+${e.diff.toFixed(2)}h overtime` : `${e.diff.toFixed(2)}h undertime`;
            return (
              <tr key={e.id}>
                <Td>{e.name}</Td>
                <Td mono>{e.target}h</Td>
                <Td mono>{e.logged.toFixed(2)}h</Td>
                <Td mono>{onTrack ? "—" : `${over ? "+" : ""}${e.diff.toFixed(2)}h`}</Td>
                <Td><Badge tone={tone}>{label}</Badge></Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 16 }}>
          <WeekNav weekStart={weekStart} setWeekStart={setWeekStart} />
          <Field label="Employee">
            <select style={{ ...inputStyle, width: 170 }} value={empFilter} onChange={(e) => setEmpFilter(e.target.value)}>
              <option value="all">All employees</option>
              {employees.filter((e) => e.active).map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
            </select>
          </Field>
          <Field label="Client">
            <select style={{ ...inputStyle, width: 170 }} value={clientFilter} onChange={(e) => setClientFilter(e.target.value)}>
              <option value="all">All clients</option>
              {clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
        </div>

        {rows.length === 0 ? (
          <EmptyState icon={Clock} title="No entries match these filters" />
        ) : (
          <TableShell headers={["Date", "Employee", "Client", "Category", "Hours", "Note", "Status", ""]}>
            {rows.map((e) => {
              const locked = isRowLocked(e);
              return (
                <tr key={e.id}>
                  <Td>{fmtDow(fromKey(e.date))} {fmtShort(fromKey(e.date))}</Td>
                  <Td>{empName(e.employeeId)}</Td>
                  <Td>{clientName(e.clientId)}</Td>
                  <Td>{CATEGORY_LABELS[e.category] || "Other"}</Td>
                  <Td mono>
                    {editingId === e.id ? (
                      <input type="number" min="0.25" step="0.25" autoFocus value={editHours}
                        onChange={(ev) => setEditHours(ev.target.value)}
                        style={{ ...inputStyle, width: 70, padding: "3px 7px", fontFamily: mono }} />
                    ) : e.hours.toFixed(2) + "h"}
                  </Td>
                  <Td>
                    {editingId === e.id ? (
                      <input type="text" placeholder="Note" value={editNote}
                        onChange={(ev) => setEditNote(ev.target.value)}
                        style={{ ...inputStyle, width: 160, padding: "3px 7px" }} />
                    ) : (e.note || "—")}
                  </Td>
                  <Td>{locked ? <Badge tone="warn">Submitted</Badge> : <Badge tone="neutral">Open</Badge>}</Td>
                  <Td right>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      {editingId === e.id ? (
                        <>
                          <button onClick={() => saveEdit(e.id)} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                          <button onClick={() => setEditingId(null)} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                        </>
                      ) : (
                        <>
                          <button onClick={() => { setEditingId(e.id); setEditHours(String(e.hours)); setEditNote(e.note || ""); }} style={iconBtnStyle}>
                            <Pencil size={13} color={C.inkMuted} />
                          </button>
                          <button onClick={() => removeEntry(e.id)} style={iconBtnStyle}>
                            <Trash2 size={13} color={C.inkMuted} />
                          </button>
                          <button onClick={() => toggleRowLock(e)} style={iconBtnStyle} title={locked ? "Reopen week" : "Lock week"}>
                            {locked ? <Unlock size={13} color={C.inkMuted} /> : <Lock size={13} color={C.inkMuted} />}
                          </button>
                        </>
                      )}
                    </div>
                  </Td>
                </tr>
              );
            })}
          </TableShell>
        )}
        <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 12, fontFamily: sans, fontSize: 13, fontWeight: 700 }}>
          Total:&nbsp;<span style={{ fontFamily: mono, color: C.accent }}>{total.toFixed(2)}h</span>
        </div>
      </Panel>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Clients                                                          */
/* ---------------------------------------------------------------------- */

function AdminClients({ clients, refetchClients, entries }) {
  const [modal, setModal] = useState(null); // null | "new" | client object
  const [name, setName] = useState("");
  const [allocAccounting, setAllocAccounting] = useState("");
  const [allocTax, setAllocTax] = useState("");
  const [allocPayroll, setAllocPayroll] = useState("");
  const [allocOther, setAllocOther] = useState("");
  const [fixedFee, setFixedFee] = useState("");
  const [feeEffectiveDate, setFeeEffectiveDate] = useState(toKey(TODAY));
  const [endDate, setEndDate] = useState("");
  const [query, setQuery] = useState("");
  const [editingHistoryId, setEditingHistoryId] = useState(null);
  const [histForm, setHistForm] = useState(null);
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);
  const [extraFeeClientId, setExtraFeeClientId] = useState(null);
  const [extraFeeMonth, setExtraFeeMonth] = useState("");
  const [extraFeeAmount, setExtraFeeAmount] = useState("");
  const [extraFeeNote, setExtraFeeNote] = useState("");

  function resetAllocFields(c) {
    const a = (c && c.allocation) || {};
    setAllocAccounting(c ? String(a.accounting || "") : "");
    setAllocTax(c ? String(a.tax || "") : "");
    setAllocPayroll(c ? String(a.payroll || "") : "");
    setAllocOther(c ? String(a.other || "") : "");
  }
  function openNew() {
    setModal("new"); setName(""); setFixedFee(""); setFeeEffectiveDate(toKey(TODAY)); setEndDate(""); setEditingHistoryId(null);
    resetAllocFields(null);
  }
  function openEdit(c) {
    setModal(c); setName(c.name); setFixedFee(c.fixedFee ?? ""); setFeeEffectiveDate(toKey(TODAY)); setEndDate(c.endDate || ""); setEditingHistoryId(null);
    resetAllocFields(c);
  }

  async function save() {
    if (!name.trim()) return;
    const fee = fixedFee === "" ? null : Number(fixedFee);
    const allocation = {
      accounting: Number(allocAccounting) || 0,
      tax: Number(allocTax) || 0,
      payroll: Number(allocPayroll) || 0,
      other: Number(allocOther) || 0,
    };
    if (modal === "new") {
      const { data, error } = await supabase.from("clients").insert({ name: name.trim(), fixed_fee: fee, end_date: endDate || null }).select().single();
      if (error || !data) {
        window.alert(`Couldn't create client: ${error ? error.message : "unknown error"}`);
        return;
      }
      const { error: feeError } = await supabase.from("client_fee_history").insert({
        client_id: data.id, effective_date: feeEffectiveDate || toKey(TODAY), fixed_fee: fee,
        alloc_accounting: allocation.accounting, alloc_tax: allocation.tax, alloc_payroll: allocation.payroll, alloc_other: allocation.other,
      });
      if (feeError) {
        window.alert(`Client was created, but saving its fee terms failed: ${feeError.message}`);
      }
    } else {
      const { error } = await supabase.from("clients").update({ name: name.trim(), fixed_fee: fee, end_date: endDate || null }).eq("id", modal.id);
      if (error) {
        window.alert(`Couldn't update client: ${error.message}`);
        return;
      }
      const prevFee = { fixedFee: modal.fixedFee ?? null, allocation: modal.allocation || { accounting: 0, tax: 0, payroll: 0, other: 0 } };
      const changed = JSON.stringify(prevFee) !== JSON.stringify({ fixedFee: fee, allocation });
      if (changed) {
        const { error: feeError } = await supabase.from("client_fee_history").insert({
          client_id: modal.id, effective_date: feeEffectiveDate || toKey(TODAY), fixed_fee: fee,
          alloc_accounting: allocation.accounting, alloc_tax: allocation.tax, alloc_payroll: allocation.payroll, alloc_other: allocation.other,
        });
        if (feeError) {
          window.alert(`Client was updated, but saving the new fee terms failed: ${feeError.message}`);
        }
      }
    }
    await refetchClients();
    setModal(null);
  }
  async function toggleActive(c) {
    await supabase.from("clients").update({ active: !c.active }).eq("id", c.id);
    refetchClients();
  }
  async function deleteClient(c) {
    if (!window.confirm(`Permanently delete "${c.name}"? This cannot be undone.`)) return;
    const { error } = await supabase.from("clients").delete().eq("id", c.id);
    if (error) {
      if (error.code === "23503") {
        window.alert(`Can't delete "${c.name}" — it still has logged hours on record. Use Deactivate instead, or remove its time entries first.`);
      } else {
        window.alert(`Couldn't delete "${c.name}": ${error.message}`);
      }
      return;
    }
    refetchClients();
  }

  function startEditHistory(h) {
    setEditingHistoryId(h.id);
    const a = h.allocation || {};
    setHistForm({
      effectiveDate: h.effectiveDate, fixedFee: h.fixedFee ?? "",
      accounting: a.accounting ?? "", tax: a.tax ?? "", payroll: a.payroll ?? "", other: a.other ?? "",
    });
  }
  function cancelEditHistory() { setEditingHistoryId(null); setHistForm(null); }
  async function saveHistoryEdit() {
    await supabase.from("client_fee_history").update({
      effective_date: histForm.effectiveDate,
      fixed_fee: histForm.fixedFee === "" ? null : Number(histForm.fixedFee),
      alloc_accounting: Number(histForm.accounting) || 0,
      alloc_tax: Number(histForm.tax) || 0,
      alloc_payroll: Number(histForm.payroll) || 0,
      alloc_other: Number(histForm.other) || 0,
    }).eq("id", editingHistoryId);
    await refetchClients();
    cancelEditHistory();
  }
  async function deleteHistoryEntry(entryId) {
    await supabase.from("client_fee_history").delete().eq("id", entryId);
    await refetchClients();
    if (editingHistoryId === entryId) cancelEditHistory();
  }

  function openExtraFee(clientId, defaultMonth) {
    setExtraFeeClientId(clientId);
    setExtraFeeMonth(defaultMonth);
    setExtraFeeAmount("");
    setExtraFeeNote("");
  }
  async function saveExtraFee() {
    if (!extraFeeMonth || !extraFeeAmount || Number(extraFeeAmount) <= 0) return;
    await supabase.from("client_extra_fees").insert({ client_id: extraFeeClientId, month: extraFeeMonth, amount: Number(extraFeeAmount), note: extraFeeNote.trim() || null });
    await refetchClients();
    setExtraFeeClientId(null);
  }
  async function deleteExtraFee(entryId) {
    await supabase.from("client_extra_fees").delete().eq("id", entryId);
    refetchClients();
  }

  const [viewRangeStart, viewRangeEnd] = periodRange(viewPeriod, viewAnchor);
  const periodEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= viewRangeStart && d <= viewRangeEnd; });
  const filtered = clients.filter((c) => c.name.toLowerCase().includes(query.toLowerCase()));
  const liveClient = modal && modal !== "new" ? clients.find((c) => c.id === modal.id) : null;

  return (
    <div>
      <PageHeader title="Clients" sub="Manage the clients your team logs hours against, and their billing terms."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={() => exportToExcel(`clients-${toKey(TODAY)}.xlsx`, [{
              name: "Clients",
              rows: clients.map((c) => ({
                Name: c.name, Status: c.active ? "Active" : "Inactive",
                "Fixed fee (€/mo)": c.fixedFee ?? "", "Accounting alloc. (h)": c.allocation?.accounting || 0,
                "Tax alloc. (h)": c.allocation?.tax || 0, "Payroll alloc. (h)": c.allocation?.payroll || 0,
                "Other alloc. (h)": c.allocation?.other || 0,
              })),
            }])}>Export to Excel</Btn>
            <Btn icon={Plus} onClick={openNew}>Add client</Btn>
          </div>
        } />
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Hours, allocation, and fee below reflect this reporting period — not just the latest settings.
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ marginBottom: 14, position: "relative", maxWidth: 260 }}>
          <Search size={14} color={C.inkFaint} style={{ position: "absolute", left: 9, top: 9 }} />
          <input placeholder="Search clients…" value={query} onChange={(e) => setQuery(e.target.value)}
            style={{ ...inputStyle, width: "100%", paddingLeft: 28 }} />
        </div>
        <TableShell headers={["Client", "Hours", "Revenue", "Allocation", "Fixed fee", "Extra fees (period)", "Status", ""]}>
          {filtered.map((c) => {
            const clientPeriodEntries = periodEntries.filter((e) => e.clientId === c.id);
            const hrs = clientPeriodEntries.reduce((s, e) => s + e.hours, 0);
            const { revenue, allocation } = periodClientMetrics(c, clientPeriodEntries, viewRangeStart, viewRangeEnd);
            const feeThen = feeAsOf(c, toKey(viewRangeStart));
            const a = feeThen.allocation || {};
            const allocTitle = CATEGORIES.map((cat) => `${CATEGORY_LABELS[cat]} ${a[cat] || 0}h`).join(" · ");
            const extraThisPeriod = extraFeesAccruedForRange(c, viewRangeStart, viewRangeEnd);
            const defaultMonth = `${viewRangeStart.getFullYear()}-${pad(viewRangeStart.getMonth() + 1)}`;
            return (
              <tr key={c.id} style={{ opacity: c.active ? 1 : 0.55 }}>
                <Td>{c.name}</Td>
                <Td mono>{hrs.toFixed(2)}h</Td>
                <Td mono>{fmtEur(revenue)}</Td>
                <Td mono title={allocTitle}>{allocation ? allocation.toFixed(1) + "h" : "—"}</Td>
                <Td mono>{feeThen.fixedFee ? fmtEur(feeThen.fixedFee) + "/mo" : "—"}</Td>
                <Td mono>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span>{extraThisPeriod > 0 ? fmtEur(extraThisPeriod) : "—"}</span>
                    <button onClick={() => openExtraFee(c.id, defaultMonth)} title="Add extra fee"
                      style={{ ...iconBtnStyle, border: `1px solid ${C.border}`, borderRadius: 5, padding: "2px 5px" }}>
                      <Plus size={11} color={C.accent} />
                    </button>
                  </div>
                </Td>
                <Td>
                  {c.endDate && c.endDate <= toKey(TODAY)
                    ? <Badge tone="warn">Ended {c.endDate}</Badge>
                    : c.active ? <Badge tone="accent">Active</Badge> : <Badge tone="neutral">Inactive</Badge>}
                </Td>
                <Td right>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openEdit(c)}>Edit</Btn>
                    <Btn size="sm" variant={c.active ? "danger" : "secondary"} onClick={() => toggleActive(c)}>
                      {c.active ? "Deactivate" : "Reactivate"}
                    </Btn>
                    <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteClient(c)}>Delete</Btn>
                  </div>
                </Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      <Panel>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Category workload</div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Allocated vs. actual hours per category, for this period.</div>
        <TableShell headers={["Client", ...CATEGORIES.flatMap((c) => [`${CATEGORY_LABELS[c]} alloc.`, `${CATEGORY_LABELS[c]} actual`])]}>
          {filtered.filter((c) => c.active).map((c) => {
            const clientPeriodEntries = periodEntries.filter((e) => e.clientId === c.id);
            const { allocationByCategory, actualByCategory } = periodClientMetrics(c, clientPeriodEntries, viewRangeStart, viewRangeEnd);
            return (
              <tr key={c.id}>
                <Td>{c.name}</Td>
                {CATEGORIES.map((cat) => (
                  <React.Fragment key={cat}>
                    <Td mono>{allocationByCategory[cat] > 0 ? allocationByCategory[cat].toFixed(1) + "h" : "—"}</Td>
                    <Td mono style={{ color: actualByCategory[cat] > allocationByCategory[cat] && allocationByCategory[cat] > 0 ? C.warn : undefined }}>
                      {actualByCategory[cat] > 0 ? actualByCategory[cat].toFixed(1) + "h" : "—"}
                    </Td>
                  </React.Fragment>
                ))}
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      {modal && (
        <Modal title={modal === "new" ? "Add client" : "Edit client"} onClose={() => setModal(null)} width={560}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Client name">
              <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Acme Retail" />
            </Field>
            <div>
              <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Agreed monthly allocation, by category</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Accounting hours">
                  <input type="number" min="0" style={inputStyle} value={allocAccounting} onChange={(e) => setAllocAccounting(e.target.value)} placeholder="e.g. 20" />
                </Field>
                <Field label="Tax hours">
                  <input type="number" min="0" style={inputStyle} value={allocTax} onChange={(e) => setAllocTax(e.target.value)} placeholder="e.g. 15" />
                </Field>
                <Field label="Payroll hours">
                  <input type="number" min="0" style={inputStyle} value={allocPayroll} onChange={(e) => setAllocPayroll(e.target.value)} placeholder="e.g. 5" />
                </Field>
                <Field label="Other hours">
                  <input type="number" min="0" style={inputStyle} value={allocOther} onChange={(e) => setAllocOther(e.target.value)} placeholder="e.g. 5" />
                </Field>
              </div>
            </div>
            <Field label="Fixed monthly fee — retainer (optional)">
              <input type="number" min="0" style={inputStyle} value={fixedFee}
                onChange={(e) => setFixedFee(e.target.value)} placeholder="e.g. 1200" />
            </Field>
            <Field label="Effective from">
              <input type="date" style={inputStyle} value={feeEffectiveDate} onChange={(e) => setFeeEffectiveDate(e.target.value)} />
            </Field>
            <Field label="End date — client relationship ended (optional)">
              <input type="date" style={inputStyle} value={endDate} onChange={(e) => setEndDate(e.target.value)} />
            </Field>
            <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, lineHeight: 1.5 }}>
              Revenue is the fixed fee (this recurring rate) plus any extra fees added below for
              specific months. The category allocation is for tracking workload only (allocated
              vs. actual), set per client on the Employees' timesheets. Saving with different
              numbers above adds a new row below, effective from that date — it never overwrites
              past periods.
            </div>
            <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink }}>Extra fees by month</div>
                {liveClient && (
                  <Btn size="sm" variant="secondary" icon={Plus} onClick={() => openExtraFee(liveClient.id, toKey(TODAY).slice(0, 7))}>Add</Btn>
                )}
              </div>
              {liveClient && liveClient.extraFeeEntries && liveClient.extraFeeEntries.length > 0 ? (
                <div className="scrollbar-thin" style={{ maxHeight: 140, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12 }}>
                    <thead>
                      <tr>
                        {["Month", "Amount", "Note", ""].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "3px 6px", fontSize: 10.5, fontWeight: 600, color: C.inkFaint, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[...liveClient.extraFeeEntries].sort((a, b) => b.month.localeCompare(a.month)).map((x) => (
                        <tr key={x.id}>
                          <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{x.month}</td>
                          <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{fmtEur(x.amount)}</td>
                          <td style={{ padding: "3px 6px", color: C.ink }}>{x.note || "—"}</td>
                          <td style={{ padding: "3px 6px" }}>
                            <button onClick={() => deleteExtraFee(x.id)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint }}>No extra fees added yet.</div>
              )}
            </div>
            {liveClient && liveClient.feeHistory && liveClient.feeHistory.length > 0 && (
              <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Fee history</div>
                <div className="scrollbar-thin" style={{ maxHeight: 220, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12 }}>
                    <thead>
                      <tr>
                        {["From", "To", "Alloc. total", "Fixed fee", ""].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "3px 6px", fontSize: 10.5, fontWeight: 600, color: C.inkFaint, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[...withPeriods(liveClient.feeHistory)].reverse().map((h) => {
                        const a = h.allocation || {};
                        const allocTotal = CATEGORIES.reduce((s, c) => s + (a[c] || 0), 0);
                        const allocTitle = CATEGORIES.map((cat) => `${CATEGORY_LABELS[cat]} ${a[cat] || 0}h`).join(" · ");
                        const isEditing = editingHistoryId === h.id;
                        return isEditing ? (
                          <tr key={h.id}>
                            <td colSpan={5} style={{ padding: "6px" }}>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                                <input type="date" style={{ ...inputStyle, width: 130, padding: "4px 6px" }} value={histForm.effectiveDate} onChange={(ev) => setHistForm((f) => ({ ...f, effectiveDate: ev.target.value }))} />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.accounting} onChange={(ev) => setHistForm((f) => ({ ...f, accounting: ev.target.value }))} placeholder="Acct h" />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.tax} onChange={(ev) => setHistForm((f) => ({ ...f, tax: ev.target.value }))} placeholder="Tax h" />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.payroll} onChange={(ev) => setHistForm((f) => ({ ...f, payroll: ev.target.value }))} placeholder="Payroll h" />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.other} onChange={(ev) => setHistForm((f) => ({ ...f, other: ev.target.value }))} placeholder="Other h" />
                                <input type="number" style={{ ...inputStyle, width: 80, padding: "4px 6px" }} value={histForm.fixedFee} onChange={(ev) => setHistForm((f) => ({ ...f, fixedFee: ev.target.value }))} placeholder="Fixed fee" />
                                <button onClick={() => saveHistoryEdit()} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                                <button onClick={cancelEditHistory} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={h.id}>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.from}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.to}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }} title={allocTitle}>{allocTotal > 0 ? allocTotal + "h" : "—"}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.fixedFee ? fmtEur(h.fixedFee) : "—"}</td>
                            <td style={{ padding: "3px 6px" }}>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button onClick={() => startEditHistory(h)} style={iconBtnStyle}><Pencil size={12} color={C.inkMuted} /></button>
                                <button onClick={() => deleteHistoryEntry(h.id)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setModal(null)}>Cancel</Btn>
              <Btn onClick={save}>Save client</Btn>
            </div>
          </div>
        </Modal>
      )}

      {extraFeeClientId && (
        <Modal title={`Add extra fee — ${clients.find((c) => c.id === extraFeeClientId)?.name || ""}`} onClose={() => setExtraFeeClientId(null)} width={360}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Month">
              <input type="month" style={inputStyle} value={extraFeeMonth} onChange={(e) => setExtraFeeMonth(e.target.value)} />
            </Field>
            <Field label="Amount">
              <input type="number" min="0" style={inputStyle} value={extraFeeAmount} onChange={(e) => setExtraFeeAmount(e.target.value)} placeholder="e.g. 150" autoFocus />
            </Field>
            <Field label="Note (optional)">
              <input style={inputStyle} value={extraFeeNote} onChange={(e) => setExtraFeeNote(e.target.value)} placeholder="e.g. Extra VAT filing" />
            </Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setExtraFeeClientId(null)}>Cancel</Btn>
              <Btn onClick={saveExtraFee}>Add fee</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Employees                                                        */
/* ---------------------------------------------------------------------- */

function AdminEmployees({ employees, refetchEmployees, entries }) {
  const [modal, setModal] = useState(null); // null | employee object (accounts are created via self sign-up)
  const [name, setName] = useState("");
  const [title, setTitle] = useState("");
  const [role, setRole] = useState("employee");
  const [weeklyHours, setWeeklyHours] = useState("40");
  const [annualLeaveDays, setAnnualLeaveDays] = useState("25");
  const [grossSalary, setGrossSalary] = useState("");
  const [socialSecurity, setSocialSecurity] = useState("");
  const [ticketRestaurant, setTicketRestaurant] = useState("");
  const [insurance, setInsurance] = useState("");
  const [otherCost, setOtherCost] = useState("");
  const [costEffectiveDate, setCostEffectiveDate] = useState(toKey(TODAY));
  const [editingHistoryId, setEditingHistoryId] = useState(null);
  const [histForm, setHistForm] = useState(null);
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);

  function resetCostFields(e) {
    setGrossSalary(e ? String(e.grossSalary || "") : "");
    setSocialSecurity(e ? String(e.socialSecurity || "") : "");
    setTicketRestaurant(e ? String(e.ticketRestaurant || "") : "");
    setInsurance(e ? String(e.insurance || "") : "");
    setOtherCost(e ? String(e.otherCost || "") : "");
    setCostEffectiveDate(toKey(TODAY));
    setEditingHistoryId(null);
  }
  function openEdit(e) {
    setModal(e); setName(e.name); setTitle(e.title || ""); setRole(e.role);
    setWeeklyHours(String(e.weeklyHours || 40)); setAnnualLeaveDays(String(e.annualLeaveDays || 25));
    resetCostFields(e);
  }

  async function save() {
    if (!modal || !name.trim()) return;
    const hrs = Number(weeklyHours) || 40;
    const leaveDays = Number(annualLeaveDays) || 25;
    const costFields = {
      grossSalary: Number(grossSalary) || 0,
      socialSecurity: Number(socialSecurity) || 0,
      ticketRestaurant: Number(ticketRestaurant) || 0,
      insurance: Number(insurance) || 0,
      otherCost: Number(otherCost) || 0,
    };
    await supabase.from("profiles").update({
      name: name.trim(), title: title.trim() || null, role,
      weekly_hours: hrs, annual_leave_days: leaveDays,
    }).eq("id", modal.id);

    const prevCost = { grossSalary: modal.grossSalary || 0, socialSecurity: modal.socialSecurity || 0, ticketRestaurant: modal.ticketRestaurant || 0, insurance: modal.insurance || 0, otherCost: modal.otherCost || 0 };
    const changed = JSON.stringify(prevCost) !== JSON.stringify(costFields);
    if (changed) {
      await supabase.from("employee_cost_history").insert({
        employee_id: modal.id, effective_date: costEffectiveDate || toKey(TODAY),
        gross_salary: costFields.grossSalary, social_security: costFields.socialSecurity,
        ticket_restaurant: costFields.ticketRestaurant, insurance: costFields.insurance, other_cost: costFields.otherCost,
      });
    }
    await refetchEmployees();
    setModal(null);
  }
  async function toggleActive(e) {
    await supabase.from("profiles").update({ active: !e.active }).eq("id", e.id);
    refetchEmployees();
  }
  async function deleteEmployee(e) {
    if (!window.confirm(`Permanently delete "${e.name}"? This removes their HoursLedger record but does not revoke their login — this cannot be undone.`)) return;
    const { error } = await supabase.from("profiles").delete().eq("id", e.id);
    if (error) {
      if (error.code === "23503") {
        window.alert(`Can't delete "${e.name}" — they still have logged hours or leave requests on record. Use Deactivate instead.`);
      } else {
        window.alert(`Couldn't delete "${e.name}": ${error.message}`);
      }
      return;
    }
    refetchEmployees();
  }

  function startEditHistory(h) {
    setEditingHistoryId(h.id);
    setHistForm({ effectiveDate: h.effectiveDate, grossSalary: String(h.grossSalary || ""), socialSecurity: String(h.socialSecurity || ""), ticketRestaurant: String(h.ticketRestaurant || ""), insurance: String(h.insurance || ""), otherCost: String(h.otherCost || "") });
  }
  function cancelEditHistory() { setEditingHistoryId(null); setHistForm(null); }
  async function saveHistoryEdit() {
    await supabase.from("employee_cost_history").update({
      effective_date: histForm.effectiveDate,
      gross_salary: Number(histForm.grossSalary) || 0,
      social_security: Number(histForm.socialSecurity) || 0,
      ticket_restaurant: Number(histForm.ticketRestaurant) || 0,
      insurance: Number(histForm.insurance) || 0,
      other_cost: Number(histForm.otherCost) || 0,
    }).eq("id", editingHistoryId);
    await refetchEmployees();
    cancelEditHistory();
  }
  async function deleteHistoryEntry(entryId) {
    await supabase.from("employee_cost_history").delete().eq("id", entryId);
    await refetchEmployees();
    if (editingHistoryId === entryId) cancelEditHistory();
  }

  const [viewRangeStart, viewRangeEnd] = periodRange(viewPeriod, viewAnchor);
  const periodEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= viewRangeStart && d <= viewRangeEnd; });
  const liveEmployee = modal ? employees.find((e) => e.id === modal.id) : null;

  return (
    <div>
      <PageHeader title="Employees" sub="People join by creating their own account on the sign-in screen. Set their role, contracted hours, and labor cost here."
        right={
          <Btn variant="secondary" onClick={() => exportToExcel(`employees-${toKey(TODAY)}.xlsx`, [{
            name: "Employees",
            rows: employees.map((e) => ({
              Name: e.name, Title: e.title || "", Email: e.email, Role: e.role, Status: e.active ? "Active" : "Inactive",
              "Contracted h/wk": e.weeklyHours || 40, "Annual leave (d/yr)": e.annualLeaveDays || 25,
              "Gross salary": e.grossSalary || 0, "Social security": e.socialSecurity || 0,
              "Ticket restaurant": e.ticketRestaurant || 0, "Other insurance": e.insurance || 0, "Other cost": e.otherCost || 0,
            })),
          }])}>Export to Excel</Btn>
        } />
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Hours and cost below reflect this reporting period — cost uses whatever was in effect at the time, not just the latest figure.
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>
      <Panel>
        <TableShell headers={["Name", "Title", "Email", "Role", "Contracted /wk", "Leave/yr", "Hours", "Cost /mo", "Cost /h", "Status", ""]}>
          {employees.map((e) => {
            const hrs = periodEntries.filter((t) => t.employeeId === e.id).reduce((s, t) => s + t.hours, 0);
            const costThenMonthly = monthlyLaborCostAsOf(e, toKey(viewRangeStart));
            const costThenHourly = hourlyLaborCostAsOf(e, toKey(viewRangeStart));
            return (
              <tr key={e.id} style={{ opacity: e.active ? 1 : 0.55 }}>
                <Td>{e.name}</Td>
                <Td>{e.title || "—"}</Td>
                <Td>{e.email}</Td>
                <Td><Badge tone={e.role === "admin" ? "accent" : "neutral"}>{e.role === "admin" ? "Manager" : "Employee"}</Badge></Td>
                <Td mono>{e.weeklyHours || 40}h</Td>
                <Td mono>{e.annualLeaveDays || 25}d</Td>
                <Td mono>{hrs.toFixed(2) + "h"}</Td>
                <Td mono>{fmtEur(costThenMonthly)}</Td>
                <Td mono>{fmtEur(costThenHourly)}</Td>
                <Td>{e.active ? <Badge tone="accent">Active</Badge> : <Badge tone="neutral">Inactive</Badge>}</Td>
                <Td right>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openEdit(e)}>Edit</Btn>
                    <Btn size="sm" variant={e.active ? "danger" : "secondary"} onClick={() => toggleActive(e)}>
                      {e.active ? "Deactivate" : "Reactivate"}
                    </Btn>
                    <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteEmployee(e)}>Delete</Btn>
                  </div>
                </Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      {modal && (
        <Modal title="Edit employee" onClose={() => setModal(null)} width={520}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Full name">
              <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Maria Ioannou" />
            </Field>
            <Field label="Job title">
              <input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Senior Accountant" />
            </Field>
            <Field label="Contracted hours per week">
              <input type="number" min="1" max="80" style={inputStyle} value={weeklyHours} onChange={(e) => setWeeklyHours(e.target.value)} placeholder="e.g. 40" />
            </Field>
            <Field label="Annual leave entitlement (days/year)">
              <input type="number" min="0" max="40" style={inputStyle} value={annualLeaveDays} onChange={(e) => setAnnualLeaveDays(e.target.value)} placeholder="e.g. 25" />
            </Field>
            <Field label="Role">
              <select style={inputStyle} value={role} onChange={(e) => setRole(e.target.value)}>
                <option value="employee">Employee</option>
                <option value="admin">Manager / Admin</option>
              </select>
            </Field>
            <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12, marginTop: 2 }}>
              <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 10 }}>Record a new labor cost change</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Gross salary">
                  <input type="number" min="0" style={inputStyle} value={grossSalary} onChange={(e) => setGrossSalary(e.target.value)} placeholder="e.g. 1400" />
                </Field>
                <Field label="Social security (employer)">
                  <input type="number" min="0" style={inputStyle} value={socialSecurity} onChange={(e) => setSocialSecurity(e.target.value)} placeholder="e.g. 320" />
                </Field>
                <Field label="Ticket restaurant">
                  <input type="number" min="0" style={inputStyle} value={ticketRestaurant} onChange={(e) => setTicketRestaurant(e.target.value)} placeholder="e.g. 88" />
                </Field>
                <Field label="Other insurance (optional)">
                  <input type="number" min="0" style={inputStyle} value={insurance} onChange={(e) => setInsurance(e.target.value)} placeholder="e.g. 0" />
                </Field>
                <Field label="Other cost">
                  <input type="number" min="0" style={inputStyle} value={otherCost} onChange={(e) => setOtherCost(e.target.value)} placeholder="Optional" />
                </Field>
                <Field label="Effective from">
                  <input type="date" style={inputStyle} value={costEffectiveDate} onChange={(e) => setCostEffectiveDate(e.target.value)} />
                </Field>
              </div>
              <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, marginTop: 8 }}>
                Cost per hour = total monthly cost ÷ (22 working days × daily hours from the contracted weekly hours).
                Saving with different numbers above adds a new row below, effective from that date — it never overwrites past periods.
              </div>
            </div>
            {liveEmployee && liveEmployee.costHistory && liveEmployee.costHistory.length > 0 && (
              <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Cost history</div>
                <div className="scrollbar-thin" style={{ maxHeight: 220, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12 }}>
                    <thead>
                      <tr>
                        {["From", "To", "Gross", "SS", "Ticket", "Ins.", "Other", "Total", ""].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "3px 6px", fontSize: 10.5, fontWeight: 600, color: C.inkFaint, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[...withPeriods(liveEmployee.costHistory)].reverse().map((h) => {
                        const total = (h.grossSalary || 0) + (h.socialSecurity || 0) + (h.ticketRestaurant || 0) + (h.insurance || 0) + (h.otherCost || 0);
                        const isEditing = editingHistoryId === h.id;
                        return isEditing ? (
                          <tr key={h.id}>
                            <td colSpan={9} style={{ padding: "6px" }}>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                                <input type="date" style={{ ...inputStyle, width: 130, padding: "4px 6px" }} value={histForm.effectiveDate} onChange={(ev) => setHistForm((f) => ({ ...f, effectiveDate: ev.target.value }))} />
                                <input type="number" style={{ ...inputStyle, width: 70, padding: "4px 6px" }} value={histForm.grossSalary} onChange={(ev) => setHistForm((f) => ({ ...f, grossSalary: ev.target.value }))} placeholder="Gross" />
                                <input type="number" style={{ ...inputStyle, width: 70, padding: "4px 6px" }} value={histForm.socialSecurity} onChange={(ev) => setHistForm((f) => ({ ...f, socialSecurity: ev.target.value }))} placeholder="SS" />
                                <input type="number" style={{ ...inputStyle, width: 70, padding: "4px 6px" }} value={histForm.ticketRestaurant} onChange={(ev) => setHistForm((f) => ({ ...f, ticketRestaurant: ev.target.value }))} placeholder="Ticket" />
                                <input type="number" style={{ ...inputStyle, width: 60, padding: "4px 6px" }} value={histForm.insurance} onChange={(ev) => setHistForm((f) => ({ ...f, insurance: ev.target.value }))} placeholder="Ins." />
                                <input type="number" style={{ ...inputStyle, width: 60, padding: "4px 6px" }} value={histForm.otherCost} onChange={(ev) => setHistForm((f) => ({ ...f, otherCost: ev.target.value }))} placeholder="Other" />
                                <button onClick={() => saveHistoryEdit()} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                                <button onClick={cancelEditHistory} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={h.id}>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.from}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.to}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.grossSalary || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.socialSecurity || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.ticketRestaurant || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.insurance || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.otherCost || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink, fontWeight: 600 }}>{fmtEur(total)}</td>
                            <td style={{ padding: "3px 6px" }}>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button onClick={() => startEditHistory(h)} style={iconBtnStyle}><Pencil size={12} color={C.inkMuted} /></button>
                                <button onClick={() => deleteHistoryEntry(h.id)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setModal(null)}>Cancel</Btn>
              <Btn onClick={save}>Save employee</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Root                                                                     */
/* ---------------------------------------------------------------------- */

export default function App() {
  const [session, setSession] = useState(undefined);
  const [view, setView] = useState(null);

  useEffect(() => {
    if (!supabaseConfigured) { setSession(null); return; }
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, sess) => setSession(sess));
    return () => sub.subscription.unsubscribe();
  }, []);

  const [profiles, refetchProfilesRaw, profilesLoading] = useTable("profiles", mapProfile, "name");
  const [costHistoryRows, refetchCostHistoryRaw, costHistoryLoading] = useTable("employee_cost_history", mapCostHistory);
  const [clientRows, refetchClientsRaw, clientsLoading] = useTable("clients", mapClientRow, "name");
  const [feeHistoryRows, refetchFeeHistoryRaw, feeHistoryLoading] = useTable("client_fee_history", mapFeeHistory);
  const [extraFeeRows, refetchExtraFeesRaw, extraFeesLoading] = useTable("client_extra_fees", mapExtraFee);
  const [entries, refetchEntries, entriesLoading] = useTable("time_entries", mapEntry);
  const [lockedWeeks, refetchLockedWeeks, lockedLoading] = useTable("locked_weeks", mapLockedWeek);
  const [leaveRequests, refetchLeaveRequests, leaveLoading] = useTable("leave_requests", mapLeaveRequest);

  const refetchEmployees = useCallback(async () => { await Promise.all([refetchProfilesRaw(), refetchCostHistoryRaw()]); }, [refetchProfilesRaw, refetchCostHistoryRaw]);
  const refetchClients = useCallback(async () => { await Promise.all([refetchClientsRaw(), refetchFeeHistoryRaw(), refetchExtraFeesRaw()]); }, [refetchClientsRaw, refetchFeeHistoryRaw, refetchExtraFeesRaw]);

  const employees = React.useMemo(() => {
    return profiles.map((p) => {
      const history = costHistoryRows.filter((h) => h.employeeId === p.id);
      const latest = latestFromHistory(history, ["grossSalary", "socialSecurity", "ticketRestaurant", "insurance", "otherCost"])
        || { grossSalary: 0, socialSecurity: 0, ticketRestaurant: 0, insurance: 0, otherCost: 0 };
      return { ...p, ...latest, costHistory: history };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [profiles, costHistoryRows]);

  const clients = React.useMemo(() => {
    return clientRows.map((c) => {
      const history = feeHistoryRows.filter((h) => h.clientId === c.id);
      const latest = latestFromHistory(history, ["fixedFee", "allocation"])
        || { fixedFee: c.fixedFee, allocation: { accounting: 0, tax: 0, payroll: 0, other: 0 } };
      const extraFeeEntries = extraFeeRows.filter((x) => x.clientId === c.id);
      return { ...c, ...latest, feeHistory: history, extraFeeEntries };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [clientRows, feeHistoryRows, extraFeeRows]);

  const user = session ? employees.find((p) => p.id === session.user.id) : null;

  useEffect(() => {
    if (user && view === null) setView(user.role === "admin" ? "dashboard" : "timesheet");
  }, [user, view]);

  useEffect(() => {
    if (session && !profilesLoading && !user) {
      // Profile row not created yet right after sign-up — try again shortly.
      const t = setTimeout(() => refetchProfilesRaw(), 1200);
      return () => clearTimeout(t);
    }
  }, [session, profilesLoading, user, refetchProfilesRaw]);

  if (!supabaseConfigured) return <SetupScreen />;
  if (session === undefined) return <LoadingScreen />;
  if (!session) return <AuthScreen />;
  if (profilesLoading || !user) return <LoadingScreen text="Setting up your account…" />;
  if (!user.active) return <LoadingScreen text="This account has been deactivated. Contact your manager." />;
  if (costHistoryLoading || clientsLoading || feeHistoryLoading || extraFeesLoading || entriesLoading || lockedLoading || leaveLoading || view === null) return <LoadingScreen />;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, fontFamily: sans }}>
      <FontImport />
      <Sidebar user={user} view={view} setView={setView} />
      <div style={{ flex: 1, padding: "28px 32px", minWidth: 0 }}>
        {view === "timesheet" && (
          <EmployeeTimesheet user={user} clients={clients} entries={entries} refetchEntries={refetchEntries} lockedWeeks={lockedWeeks} refetchLockedWeeks={refetchLockedWeeks} />
        )}
        {view === "timeoff" && (
          <TimeOff user={user} employees={employees} leaveRequests={leaveRequests} refetchLeaveRequests={refetchLeaveRequests} />
        )}
        {user.role === "admin" && view === "leaveapprovals" && (
          <LeaveApprovals user={user} employees={employees} leaveRequests={leaveRequests} refetchLeaveRequests={refetchLeaveRequests} />
        )}
        {view === "calendar" && (
          <TeamCalendar employees={employees} leaveRequests={leaveRequests} />
        )}
        {user.role === "admin" && view === "dashboard" && (
          <AdminDashboard employees={employees} clients={clients} entries={entries} />
        )}
        {user.role === "admin" && view === "timesheets" && (
          <AdminTimesheets employees={employees} clients={clients} entries={entries} refetchEntries={refetchEntries} lockedWeeks={lockedWeeks} refetchLockedWeeks={refetchLockedWeeks} />
        )}
        {user.role === "admin" && view === "clients" && (
          <AdminClients clients={clients} refetchClients={refetchClients} entries={entries} />
        )}
        {user.role === "admin" && view === "employees" && (
          <AdminEmployees employees={employees} refetchEmployees={refetchEmployees} entries={entries} />
        )}
      </div>
    </div>
  );
}
