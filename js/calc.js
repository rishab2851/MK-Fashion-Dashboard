// Business rules for M K Fashion vendor payables.
// Every figure on every screen comes from these functions, so the logic lives in one place.

export const DEPTS = { GREIGE: 'GREIGE', PROCESS: 'PROCESS' };
export const DEPT_LABEL = { GREIGE: 'Greige', PROCESS: 'Process House' };
export const STATUS = {
  URGENT: 'NEED URGENT ATTENTION',
  PENDING: 'PENDING',
  PARTIAL: 'PARTIALLY PAID',
  LATE: 'PAID AFTER DUE DATE',
  ONTIME: 'PAID ON TIME',
};
export const STATUS_ORDER = [STATUS.URGENT, STATUS.PENDING, STATUS.PARTIAL, STATUS.LATE, STATUS.ONTIME];
export const STATUS_CLASS = {
  [STATUS.URGENT]: 'st-urgent', [STATUS.PENDING]: 'st-pending', [STATUS.PARTIAL]: 'st-partial',
  [STATUS.LATE]: 'st-late', [STATUS.ONTIME]: 'st-ontime',
};

export const DEFAULT_SETTINGS = {
  company: 'M K FASHION',
  interestRate: 0.12,      // greige late interest, per annum
  daysInYear: 365,
  processGraceDays: 5,     // process house bills turn urgent only after this many days past due
  greigeTerms: 65,
  processTerms: 20,
};

// ---------- dates (stored as YYYY-MM-DD strings, compared as whole days) ----------
export function dayNum(s) {
  if (!s) return null;
  const [y, m, d] = s.split('-').map(Number);
  return Math.round(Date.UTC(y, m - 1, d) / 86400000);
}
export function fromDayNum(n) {
  const dt = new Date(n * 86400000);
  return dt.toISOString().slice(0, 10);
}
export function todayStr() {
  const t = new Date();
  const p = (x) => String(x).padStart(2, '0');
  return `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())}`;
}
export function addDays(s, n) { return fromDayNum(dayNum(s) + Number(n || 0)); }

// Indian financial year label for a date, e.g. 2026-11-19 -> "2026-27"
export function fyOf(s) {
  if (!s) return '';
  const [y, m] = s.split('-').map(Number);
  const start = m >= 4 ? y : y - 1;
  return `${start}-${String((start + 1) % 100).padStart(2, '0')}`;
}
export function currentFY() { return fyOf(todayStr()); }

const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);
const r2 = (v) => Math.round(v * 100) / 100;

// ---------- one bill ----------
export function computeBill(b, settings = DEFAULT_SETTINGS, today = todayStr()) {
  const s = { ...DEFAULT_SETTINGS, ...settings };
  const terms = num(b.terms);
  const due = b.invDate ? addDays(b.invDate, terms) : null;
  const pays = Array.isArray(b.payments) ? b.payments : [];
  const paid = r2(pays.reduce((a, p) => a + num(p.amount), 0));
  const tds = r2(pays.reduce((a, p) => a + num(p.tds), 0));
  const dn = r2(num(b.dnAmt));
  const inv = r2(num(b.invAmt));
  const settled = r2(paid + tds + dn);
  const bal = r2(inv - settled);
  const lastPay = pays.filter((p) => p.date).map((p) => p.date).sort().pop() || null;

  const dueN = dayNum(due);
  const todayN = dayNum(today);
  // Positive = days left before due date, negative = days past due.
  // Open bills are measured against today; settled bills against the last payment date.
  const refN = bal > 0 || !lastPay ? todayN : dayNum(lastPay);
  const days = dueN == null ? 0 : dueN - refN;

  const grace = b.dept === DEPTS.PROCESS ? num(s.processGraceDays) : 0;
  let status;
  if (bal <= 0) status = days >= 0 ? STATUS.ONTIME : STATUS.LATE;
  else if (days < -grace) status = STATUS.URGENT;
  else if (paid > 0) status = STATUS.PARTIAL;
  else status = STATUS.PENDING;

  // Interest applies to greige only: on each instalment paid after due date,
  // and accrued on the open balance of an overdue bill.
  let intPaid = 0, intAccrued = 0;
  if (b.dept === DEPTS.GREIGE && dueN != null) {
    const k = num(s.interestRate) / (num(s.daysInYear) || 365);
    for (const p of pays) {
      if (!p.date) continue;
      const late = dayNum(p.date) - dueN;
      if (late > 0) intPaid += num(p.amount) * late * k;
    }
    if (bal > 0 && days < 0) intAccrued = bal * -days * k;
  }
  return {
    due, paid, tds, dn, settled, inv, bal, lastPay, days, status,
    intPaid: r2(intPaid), intAccrued: r2(intAccrued), interest: r2(intPaid + intAccrued),
    open: bal > 0,
  };
}

// Interest that would be payable if the open balance were paid on payDate (greige only).
export function interestIfPaidOn(b, c, payDate, settings = DEFAULT_SETTINGS) {
  const s = { ...DEFAULT_SETTINGS, ...settings };
  if (b.dept !== DEPTS.GREIGE || !c.due || c.bal <= 0) return 0;
  const late = dayNum(payDate) - dayNum(c.due);
  return late > 0 ? r2(c.bal * late * num(s.interestRate) / (num(s.daysInYear) || 365)) : 0;
}

// ---------- aggregates for the dashboard ----------
export const AGING = [
  { label: 'Overdue 90+ Days', tier: 'CRITICAL', test: (d) => d < -90, overdue: true },
  { label: 'Overdue 61 to 90 Days', tier: 'HIGH', test: (d) => d >= -90 && d < -60, overdue: true },
  { label: 'Overdue 31 to 60 Days', tier: 'HIGH', test: (d) => d >= -60 && d < -30, overdue: true },
  { label: 'Overdue 1 to 30 Days', tier: 'MODERATE', test: (d) => d >= -30 && d < 0, overdue: true },
  { label: 'Due in 0 to 7 Days', tier: 'IMMEDIATE', test: (d) => d >= 0 && d <= 7 },
  { label: 'Due in 8 to 30 Days', tier: 'UPCOMING', test: (d) => d > 7 && d <= 30 },
  { label: 'Due in 31+ Days', tier: 'SCHEDULED', test: (d) => d > 30 },
];

export function summarize(rows) {
  // rows: [{bill, c}] already filtered to one department
  const kpi = { count: rows.length, inv: 0, settled: 0, bal: 0, urgent: 0, urgentCount: 0, openCount: 0,
    interest: 0, deductions: 0 };
  const byStatus = Object.fromEntries(STATUS_ORDER.map((s) => [s, { count: 0, inv: 0, paid: 0, ded: 0, bal: 0, interest: 0 }]));
  const aging = AGING.map((a) => ({ ...a, bal: 0, count: 0, interest: 0 }));
  const vendors = new Map();
  for (const { bill, c } of rows) {
    kpi.inv += c.inv; kpi.settled += c.settled; kpi.bal += c.bal; kpi.interest += c.interest;
    kpi.deductions += c.tds + c.dn;
    if (c.status === STATUS.URGENT) { kpi.urgent += c.bal; kpi.urgentCount++; }
    if (c.bal > 0) kpi.openCount++;
    const st = byStatus[c.status];
    st.count++; st.inv += c.inv; st.paid += c.paid; st.ded += c.tds + c.dn; st.bal += c.bal; st.interest += c.interest;
    if (c.bal > 0) {
      const a = aging.find((x) => x.test(c.days));
      if (a) { a.bal += c.bal; a.count++; a.interest += c.intAccrued; }
    }
    const key = bill.vendorId || bill.vendorName;
    if (!vendors.has(key)) vendors.set(key, { vendorId: bill.vendorId, name: bill.vendorName, count: 0, inv: 0, paid: 0, ded: 0, bal: 0, urgent: 0, upcoming: 0, interest: 0 });
    const v = vendors.get(key);
    v.count++; v.inv += c.inv; v.paid += c.paid; v.ded += c.tds + c.dn; v.bal += c.bal; v.interest += c.interest;
    if (c.status === STATUS.URGENT) v.urgent += c.bal;
    else if (c.bal > 0) v.upcoming += c.bal;
  }
  const vendorList = [...vendors.values()].sort((a, b) => b.bal - a.bal || b.inv - a.inv || a.name.localeCompare(b.name));
  return { kpi, byStatus, aging, vendors: vendorList };
}

// ---------- ids ----------
export function slug(s) {
  return String(s ?? '').trim().toUpperCase().replace(/[^A-Z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'X';
}
export function normName(s) { return String(s ?? '').trim().replace(/\s+/g, ' ').toUpperCase(); }
// Deterministic bill id: the same vendor, invoice number and financial year always map to one bill,
// so re-importing the Excel updates bills instead of duplicating them.
export function billId(dept, vendorName, fy, invNo) {
  return `${dept}_${slug(vendorName)}_${fy}_${slug(invNo)}`.slice(0, 700);
}
export function vendorIdOf(name) { return slug(normName(name)); }
export function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }
