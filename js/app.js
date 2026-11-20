import { firebaseConfig, IDLE_LOCK_MINUTES, SYNC_CHECK_SECONDS } from './config.js';
import * as od from './onedrive.js';
import { createStore } from './store.js';
import {
  DEPTS, DEPT_LABEL, STATUS, STATUS_ORDER, STATUS_CLASS, DEFAULT_SETTINGS, AGING,
  computeBill, summarize, interestIfPaidOn, todayStr, currentFY, fyOf, dayNum, addDays, billId, vendorIdOf, normName, uid,
} from './calc.js';
import { parseWorkbook, planSync, removalOps, exportWorkbook } from './importer.js';
import {
  esc, rs, rs2, nf, pct, dmy, daysText, $, $$, toast, openModal, closeModal, confirmBox, formData,
  downloadText, toCSV, sha256, columnChart, stackedBars,
} from './ui.js';

// ============================================================ state
const S = {
  store: null, user: null, profile: null, started: false,
  settings: { ...DEFAULT_SETTINGS }, vendors: [], bills: [], users: [], rows: [],
  fy: localStorage.getItem('mkf_fy') || currentFY(),
  dept: localStorage.getItem('mkf_dept') || DEPTS.GREIGE,
  view: 'dashboard',
  ui: {
    [DEPTS.GREIGE]: { q: '', status: 'ALL', vendor: '', sort: 'due' },
    [DEPTS.PROCESS]: { q: '', status: 'ALL', vendor: '', sort: 'due' },
    vendors: { cat: 'ALL', q: '' },
    plan: { dept: 'ALL', win: '7', date: todayStr(), q: '', sel: new Set() },
    ledger: { vendorId: '', from: '', to: '' },
  },
  ledgerCache: {},
  sync: {}, allBills: null, syncBusy: false, syncErr: '', pending: [],
};
const role = () => S.profile?.role || 'viewer';
const canEdit = () => ['admin', 'editor'].includes(role());
const isAdmin = () => role() === 'admin';
const VCATS = { GREIGE: 'Greige Supplier', PROCESS: 'Process House', TRANSPORT: 'Transport', OTHER: 'Other' };
const MODES = ['NEFT', 'RTGS', 'IMPS', 'UPI', 'Cheque', 'Cash', 'Adjustment', 'Other'];

// ============================================================ icons & nav
const I = {
  dash: '<path d="M3 13h8V3H3zM13 21h8V11h-8zM3 21h8v-6H3zM13 3v6h8V3z"/>',
  process: '<path d="M12 3v3M12 18v3M4.2 6.2l2.1 2.1M17.7 15.7l2.1 2.1M3 12h3M18 12h3M4.2 17.8l2.1-2.1M17.7 8.3l2.1-2.1"/><circle cx="12" cy="12" r="3.5"/>',
  fabric: '<path d="M4 4h12a4 4 0 0 1 0 8H8a4 4 0 0 0 0 8h12"/><path d="M4 4v16"/>',
  users: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.8-3.5 3.4-5.5 6.5-5.5s5.7 2 6.5 5.5"/><path d="M16 4.5a3.5 3.5 0 0 1 0 7M18 14.8c2 .7 3.2 2.5 3.6 5.2"/>',
  cal: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4M8 14h3v3H8z"/>',
  book: '<path d="M4 4h11a3 3 0 0 1 3 3v13H7a3 3 0 0 1-3-3z"/><path d="M4 17a3 3 0 0 1 3-3h11"/>',
  swap: '<path d="M7 4v14M3 14l4 4 4-4M17 20V6M13 10l4-4 4 4"/>',
  gear: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
};
const icon = (k) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${I[k]}</svg>`;
const NAV = [
  { grp: 'Payables' },
  { id: 'dashboard', label: 'Vendor Payment Dashboard', icon: 'dash' },
  { id: 'process', label: 'Process House', icon: 'process', dept: DEPTS.PROCESS },
  { id: 'greige', label: 'Greige', icon: 'fabric', dept: DEPTS.GREIGE },
  { id: 'vendors', label: 'Vendor Master List', icon: 'users' },
  { grp: 'Tools' },
  { id: 'planner', label: 'Payment Planner', icon: 'cal' },
  { id: 'ledger', label: 'Vendor Ledger', icon: 'book' },
  { id: 'data', label: 'Excel Sync', icon: 'swap', need: 'edit' },
  { id: 'settings', label: 'Settings and Users', icon: 'gear', need: 'admin' },
];
const allowed = (n) => !n.need || (n.need === 'edit' ? canEdit() : isAdmin());

// ============================================================ boot & auth
const SCREENS = ['boot', 'login', 'noaccess', 'lock', 'app'];
function show(id) { SCREENS.forEach((s) => ($('#' + s).hidden = s !== id)); }

async function boot() {
  try { S.store = await createStore(firebaseConfig); }
  catch (e) {
    show('boot'); $('#boot p').textContent = 'Could not connect. Check the internet connection and reload. (' + e.message + ')'; return;
  }
  $('#btn-signin').onclick = async () => {
    $('#login-err').textContent = '';
    try { await S.store.signIn(); } catch (e) { $('#login-err').textContent = friendlyAuth(e); }
  };
  $('#btn-na-out').onclick = () => S.store.signOut();
  $('#btn-pin-out').onclick = () => signOut();
  $('#pin-form').onsubmit = onPinSubmit;
  S.store.onAuth(async (user) => {
    if (!user) { S.user = null; show('login'); return; }
    S.user = user;
    let prof = null;
    try { prof = await S.store.getProfile(user.email); } catch (e) { console.warn(e); }
    if (!prof || prof.active === false) { $('#na-email').textContent = user.email; show('noaccess'); return; }
    S.profile = prof;
    if (!prof.pinHash) return pinScreen('set');
    if (recentlyActive()) return enterApp();
    pinScreen('unlock');
  });
}
function friendlyAuth(e) {
  if (e.code === 'auth/unauthorized-domain') return 'This web address is not yet authorised in Firebase (Authentication > Settings > Authorised domains).';
  if (e.code === 'auth/popup-closed-by-user') return 'Sign in was cancelled.';
  return e.message || String(e);
}
async function signOut() { sessionStorage.removeItem('mkf_act'); await S.store.signOut(); location.hash = ''; location.reload(); }

// ---------- PIN ----------
let pinMode = 'unlock', pinTries = 0;
const pinHash = (pin) => sha256(`mkf:${S.user.email}:${pin}`);
function pinScreen(mode) {
  pinMode = mode;
  $('#pin-title').textContent = mode === 'set' ? 'Create your app PIN' : 'Enter your PIN';
  $('#pin-sub').textContent = mode === 'set' ? 'Choose a 4 to 6 digit PIN. You will need it each time the app opens or locks.' : S.user.email;
  $('#pin2').hidden = mode !== 'set'; $('#pin2').required = mode === 'set';
  $('#pin1').value = ''; $('#pin2').value = ''; $('#pin-err').textContent = '';
  $('#pin-form button[type=submit]').textContent = mode === 'set' ? 'Save PIN' : 'Unlock';
  show('lock'); setTimeout(() => $('#pin1').focus(), 50);
}
async function onPinSubmit(e) {
  e.preventDefault();
  const p1 = $('#pin1').value.trim(), p2 = $('#pin2').value.trim();
  if (!/^\d{4,6}$/.test(p1)) { $('#pin-err').textContent = 'PIN must be 4 to 6 digits.'; return; }
  if (pinMode === 'set') {
    if (p1 !== p2) { $('#pin-err').textContent = 'The two PINs do not match.'; return; }
    const h = await pinHash(p1);
    try { await S.store.setOwnPin(h); } catch (er) { $('#pin-err').textContent = 'Could not save PIN: ' + er.message; return; }
    S.profile.pinHash = h; toast('PIN saved');
    return enterApp();
  }
  if ((await pinHash(p1)) === S.profile.pinHash) { pinTries = 0; return enterApp(); }
  pinTries++;
  if (pinTries >= 5) { $('#pin-err').textContent = 'Too many wrong attempts. Signing out.'; setTimeout(signOut, 1200); return; }
  $('#pin-err').textContent = `Wrong PIN. ${5 - pinTries} attempts left.`; $('#pin1').value = '';
}

// ---------- idle lock ----------
let lastAct = Date.now();
function markActive() { lastAct = Date.now(); sessionStorage.setItem('mkf_act', JSON.stringify({ t: lastAct, u: S.user?.email })); }
function recentlyActive() {
  try { const a = JSON.parse(sessionStorage.getItem('mkf_act')); return a && a.u === S.user.email && Date.now() - a.t < IDLE_LOCK_MINUTES * 60000; } catch { return false; }
}
function startIdleWatch() {
  let th = 0;
  const bump = () => { if (Date.now() - th > 5000) { th = Date.now(); markActive(); } };
  ['pointerdown', 'keydown', 'scroll', 'touchstart'].forEach((ev) => window.addEventListener(ev, bump, { passive: true }));
  const check = () => { if (!$('#app').hidden && Date.now() - lastAct > IDLE_LOCK_MINUTES * 60000) lockNow(); };
  setInterval(check, 20000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) check(); });
}
function lockNow() { closeModal(); sessionStorage.removeItem('mkf_act'); pinScreen('unlock'); }

// ============================================================ app shell
function enterApp() {
  markActive();
  show('app');
  if (S.started) return;
  S.started = true;
  $('#demo-banner').hidden = S.store.mode !== 'demo';
  $('#mode-badge').textContent = S.store.mode === 'demo' ? 'DEMO MODE' : '';
  buildFY(); buildUser();
  $('#btn-menu').onclick = () => document.body.classList.toggle('menu-open');
  $('#side-shade').onclick = () => document.body.classList.remove('menu-open');
  window.addEventListener('hashchange', route);
  startIdleWatch();
  subscribe();
  route();
  startSync();
}
function subscribe() {
  S.store.subscribe(S.fy, {
    vendors: (v) => { S.vendors = v; changed(); },
    bills: (b) => { S.bills = b; changed(); },
    settings: (s) => { S.settings = { ...DEFAULT_SETTINGS, ...s }; changed(); },
    users: (u) => { S.users = u; changed(); },
    sync: (x) => { const prev = S.sync.syncedAt; S.sync = x || {}; chip(); if (x && x.syncedAt && x.syncedAt !== prev) maybeShowSummary(); if (S.view === 'data') render(); },
    error: (e) => toast('Sync problem: ' + (e.code || e.message), 'err'),
  }, isAdmin());
}
let chT;
function changed() {
  clearTimeout(chT);
  chT = setTimeout(() => {
    const t = todayStr();
    S.rows = S.bills.map((bill) => ({ bill, c: computeBill(bill, S.settings, t) }));
    $('#side-company').textContent = S.settings.company || 'M K FASHION';
    buildNav(); render(); chip();
  }, 40);
}
function buildFY() {
  const sel = $('#fy-select');
  const cur = Number(currentFY().slice(0, 4));
  const opts = [];
  for (let y = cur; y >= 2024; y--) opts.push(`${y}-${String((y + 1) % 100).padStart(2, '0')}`);
  if (!opts.includes(S.fy) && S.fy !== 'ALL') S.fy = opts[0];
  sel.innerHTML = opts.map((o) => `<option value="${o}">${o}</option>`).join('') + '<option value="ALL">All years</option>';
  sel.value = S.fy;
  sel.onchange = () => { S.fy = sel.value; localStorage.setItem('mkf_fy', S.fy); S.bills = []; subscribe(); };
}
function buildUser() {
  const u = S.user, btn = $('#btn-user'), pop = $('#user-pop');
  if (u.photo) btn.style.backgroundImage = `url("${u.photo}")`; else btn.textContent = (u.name || u.email)[0].toUpperCase();
  $('#side-user').innerHTML = `<div>${esc(u.name)}</div><span class="role">${esc(role())}</span>`;
  pop.innerHTML = `<div class="who"><b>${esc(u.name)}</b><br>${esc(u.email)}<br><span class="tag st-late">${esc(role().toUpperCase())}</span></div>
    <button data-a="pin">Change my PIN</button><button data-a="lock">Lock now</button><button data-a="out">Sign out</button>`;
  btn.onclick = (e) => { e.stopPropagation(); pop.hidden = !pop.hidden; };
  document.addEventListener('click', () => (pop.hidden = true));
  pop.onclick = (e) => {
    const a = e.target.dataset.a; if (!a) return;
    pop.hidden = true;
    if (a === 'pin') changePin(); if (a === 'lock') lockNow(); if (a === 'out') signOut();
  };
}
function buildNav() {
  const urgent = (d) => S.rows.filter((r) => r.bill.dept === d && r.c.status === STATUS.URGENT).length;
  $('#nav').innerHTML = NAV.filter(allowed).map((n) => {
    if (n.grp) return `<div class="grp">${n.grp}</div>`;
    const cnt = n.dept ? urgent(n.dept) : 0;
    return `<a href="#/${n.id}" class="${S.view === n.id ? 'on' : ''}">${icon(n.icon)}<span>${n.label}</span>${cnt ? `<span class="count" title="Need urgent attention">${cnt}</span>` : ''}</a>`;
  }).join('');
}
function route() {
  const [id, arg] = (location.hash.replace(/^#\/?/, '') || 'dashboard').split('/');
  const n = NAV.find((x) => x.id === id && allowed(x));
  S.view = n ? id : 'dashboard';
  if (arg && S.view === 'ledger') S.ui.ledger.vendorId = decodeURIComponent(arg);
  document.body.classList.remove('menu-open');
  buildNav(); render(); window.scrollTo(0, 0);
}
function render() {
  const n = NAV.find((x) => x.id === S.view);
  $('#page-title').textContent = n.label;
  document.title = `${n.label} | ${S.settings.company || 'M K Fashion'}`;
  (VIEWS[S.view] || VIEWS.dashboard)($('#view'));
}

// ============================================================ shared bits
const tag = (st) => `<span class="tag ${STATUS_CLASS[st]}">${esc(st)}</span>`;
const td = (l, v, cls = '') => `<td data-l="${esc(l)}" class="${cls}">${v}</td>`;
const deptRows = (d) => S.rows.filter((r) => r.bill.dept === d);
const fyNote = () => (S.fy === 'ALL' ? 'All years' : `FY ${S.fy} bills, plus older bills still open`);
const daysCell = (c) => {
  const t = daysText(c.days, c.open);
  return `<span class="${c.open && c.days < 0 ? 'neg' : ''}">${esc(t)}</span>`;
};
const vendorById = (id) => S.vendors.find((v) => v.id === id);
async function guard(fn, okMsg) {
  try { await fn(); if (okMsg) toast(okMsg); return true; }
  catch (e) { console.error(e); toast(e.code === 'permission-denied' ? 'You do not have permission for this.' : 'Error: ' + e.message, 'err'); return false; }
}

// ============================================================ VIEWS
const VIEWS = {};

// ---------------------------------------------------------------- dashboard
VIEWS.dashboard = (el) => {
  const d = S.dept, isG = d === DEPTS.GREIGE, s = S.settings;
  const rows = deptRows(d);
  const { kpi, byStatus, aging, vendors } = summarize(rows);
  const caption = isG
    ? `Showing: Greige fabric vendors (${s.greigeTerms} / 90 day credit terms | ${+(s.interestRate * 100).toFixed(2)}% p.a. late interest)`
    : `Showing: Process house vendors (${s.processTerms} day credit terms | ${s.processGraceDays} day grace)`;
  const tierCls = { CRITICAL: 't-crit', HIGH: 't-high', MODERATE: 't-mod', IMMEDIATE: 't-imm', UPCOMING: 't-up', SCHEDULED: 't-sch' };
  const intLabel = isG ? 'Interest (paid late + accrued)' : 'Total TDS + Debit notes';

  el.innerHTML = `
  <div class="dept-bar">
    <div class="seg" id="dept-seg"><button data-d="GREIGE" class="${isG ? 'on' : ''}">GREIGE</button><button data-d="PROCESS" class="${!isG ? 'on' : ''}">PROCESS HOUSE</button></div>
    <div class="dept-caption">${esc(caption)}</div>
    <div class="fine">${esc(fyNote())}</div>
  </div>
  ${!S.bills.length ? `<div class="card pad empty">No bills yet. ${canEdit() ? '<a href="#/data">Set up Excel Sync</a> to load the data from your OneDrive workbook.' : 'The data appears here once the Excel sync is set up.'}</div>` : ''}
  <div class="kpis">
    <div class="kpi"><div class="lbl">Total invoiced</div><div class="val">${rs(kpi.inv)}</div><div class="sub">${kpi.count} ${DEPT_LABEL[d]} bills</div></div>
    <div class="kpi k-settled"><div class="lbl">Total settled</div><div class="val">${rs(kpi.settled)}</div><div class="sub">${isG ? 'Paid + debit notes' : 'Paid + TDS / DN'}</div></div>
    <div class="kpi k-open"><div class="lbl">Open balance due</div><div class="val">${rs(kpi.bal)}</div><div class="sub">${kpi.openCount} unpaid bills</div></div>
    <div class="kpi k-urgent"><div class="lbl">Need urgent attention</div><div class="val">${rs(kpi.urgent)}</div><div class="sub">${kpi.urgentCount} overdue bills</div></div>
    <div class="kpi k-int"><div class="lbl">${intLabel}</div><div class="val">${rs(isG ? kpi.interest : kpi.deductions)}</div><div class="sub">${isG ? `${+(s.interestRate * 100).toFixed(2)}% p.a. on late payments and overdue bills` : 'Deducted on process bills'}</div></div>
  </div>

  <div class="card"><div class="sec-title">1. ${DEPT_LABEL[d].toUpperCase()} PAYMENT STATUS BREAKDOWN <small>Tap a status to see the bills</small></div>
   <div class="tbl-wrap"><table class="t cards"><thead><tr><th>Payment status</th><th class="num">Bills</th><th class="num">Total invoiced</th><th class="num">Amount paid</th><th class="num">${isG ? 'Debit notes' : 'TDS / DN'}</th><th class="num">Open balance</th><th class="num">% of billed</th>${isG ? '<th class="num">Interest</th>' : ''}</tr></thead>
   <tbody>${STATUS_ORDER.map((st) => { const x = byStatus[st]; return `<tr class="click" data-st="${st}">
     ${td('Status', tag(st), 'first')}${td('Bills', nf(x.count), 'num')}${td('Invoiced', rs(x.inv), 'num')}${td('Paid', rs(x.paid), 'num')}${td(isG ? 'Debit notes' : 'TDS / DN', rs(x.ded), 'num')}
     ${td('Open balance', rs(x.bal), 'num')}${td('% of billed', pct(kpi.inv ? x.inv / kpi.inv : 0), 'num')}${isG ? td('Interest', rs(x.interest), 'num') : ''}</tr>`; }).join('')}</tbody>
   <tfoot><tr>${td('', 'TOTAL DEPARTMENT', 'first')}${td('Bills', nf(kpi.count), 'num')}${td('Invoiced', rs(kpi.inv), 'num')}${td('Paid', rs(kpi.settled - kpi.deductions), 'num')}${td('Deductions', rs(kpi.deductions), 'num')}${td('Open balance', rs(kpi.bal), 'num')}${td('%', '100%', 'num')}${isG ? td('Interest', rs(kpi.interest), 'num') : ''}</tr></tfoot></table></div></div>

  <div class="grid2 section" style="margin-top:18px">
   <div class="card"><div class="sec-title">2. UNPAID AGING AND UPCOMING DUE SCHEDULE</div>
    <div class="tbl-wrap"><table class="t cards"><thead><tr><th>Aging / due window</th><th>Risk tier</th><th class="num">Open balance</th><th class="num">% of open</th><th class="num">Bills</th>${isG ? '<th class="num">Interest accrued</th>' : ''}</tr></thead>
    <tbody>${aging.map((a) => `<tr>${td('Window', esc(a.label), 'first')}${td('Risk tier', `<span class="tier ${tierCls[a.tier]}">${a.tier}</span>`)}${td('Open balance', rs(a.bal), 'num')}${td('% of open', pct(kpi.bal ? a.bal / kpi.bal : 0), 'num')}${td('Bills', nf(a.count), 'num')}${isG ? td('Interest', a.overdue ? rs(a.interest) : '', 'num') : ''}</tr>`).join('')}</tbody>
    <tfoot><tr>${td('', 'TOTAL UNPAID EXPOSURE', 'first')}${td('', 'ALL OPEN')}${td('Open balance', rs(kpi.bal), 'num')}${td('%', '100%', 'num')}${td('Bills', nf(kpi.openCount), 'num')}${isG ? td('Interest', rs(aging.reduce((x, a) => x + a.interest, 0)), 'num') : ''}</tr></tfoot></table></div></div>
   <div class="card"><div class="sec-title">Unpaid payables by aging window</div><div class="pad">
    ${columnChart(aging.map((a) => ({ label: a.label.replace('Overdue ', 'Overdue ').replace(' Days', 'd'), value: a.bal })), { colors: ['#9B1C1C', '#C81E1E', '#E8534A', '#F28E86', '#1E6B7B', '#3A8C9C', '#7FB8C4'] })}</div></div>
  </div>

  <div class="card section" style="margin-top:18px"><div class="sec-title">3. ${DEPT_LABEL[d].toUpperCase()} VENDOR WISE PAYABLES (${vendors.length} ACTIVE VENDORS, HIGHEST BALANCE FIRST) <small>Tap a vendor to open the ledger</small></div>
   <div class="pad">${stackedBars(vendors.filter((v) => v.bal > 0).slice(0, 10).map((v) => ({ label: v.name, values: [v.urgent, v.upcoming] })), [{ name: 'Urgent overdue', color: '#C81E1E' }, { name: 'Pending / upcoming', color: '#2563EB' }])}</div>
   <div class="tbl-wrap"><table class="t cards"><thead><tr><th>Vendor name</th><th class="num">Bills</th><th class="num">Total invoiced</th><th class="num">Amount paid</th><th class="num">${isG ? 'Debit notes' : 'TDS / DN'}</th><th class="num">Open balance</th><th class="num">Urgent overdue</th><th class="num">Pending / upcoming</th><th class="num">% of open</th>${isG ? '<th class="num">Interest</th>' : ''}</tr></thead>
   <tbody>${vendors.map((v) => `<tr class="click" data-v="${esc(v.vendorId)}">${td('Vendor', `<span class="vname">${esc(v.name)}</span>`, 'first')}${td('Bills', nf(v.count), 'num')}${td('Invoiced', rs(v.inv), 'num')}${td('Paid', rs(v.paid), 'num hide-m')}${td('Deductions', rs(v.ded), 'num hide-m')}${td('Open balance', `<b>${rs(v.bal)}</b>`, 'num')}${td('Urgent overdue', v.urgent ? `<span class="neg">${rs(v.urgent)}</span>` : rs(0), 'num')}${td('Pending / upcoming', rs(v.upcoming), 'num')}${td('% of open', pct(kpi.bal ? v.bal / kpi.bal : 0), 'num hide-m')}${isG ? td('Interest', rs(v.interest), 'num') : ''}</tr>`).join('') || `<tr><td colspan="10" class="empty">No bills in this view.</td></tr>`}</tbody>
   <tfoot><tr>${td('', `TOTAL ${DEPT_LABEL[d].toUpperCase()} VENDORS (${vendors.length})`, 'first')}${td('Bills', nf(kpi.count), 'num')}${td('Invoiced', rs(kpi.inv), 'num')}${td('Paid', rs(kpi.settled - kpi.deductions), 'num hide-m')}${td('Deductions', rs(kpi.deductions), 'num hide-m')}${td('Open balance', rs(kpi.bal), 'num')}${td('Urgent', rs(kpi.urgent), 'num')}${td('Pending / upcoming', rs(kpi.bal - kpi.urgent), 'num')}${td('%', '100%', 'num hide-m')}${isG ? td('Interest', rs(kpi.interest), 'num') : ''}</tr></tfoot></table></div></div>`;

  $$('#dept-seg button', el).forEach((b) => (b.onclick = () => { S.dept = b.dataset.d; localStorage.setItem('mkf_dept', S.dept); render(); }));
  $$('tr[data-st]', el).forEach((tr) => (tr.onclick = () => { S.ui[d].status = tr.dataset.st; location.hash = '#/' + (isG ? 'greige' : 'process'); }));
  $$('tr[data-v]', el).forEach((tr) => (tr.onclick = () => { location.hash = '#/ledger/' + encodeURIComponent(tr.dataset.v); }));
};

// ---------------------------------------------------------------- bill lists
function billList(dept) {
  return (el) => {
    const st = S.ui[dept], isG = dept === DEPTS.GREIGE;
    const rows = deptRows(dept);
    const vendorsHere = [...new Set(rows.map((r) => r.bill.vendorName))].sort();
    el.innerHTML = `
      <div class="toolbar">
        <input type="search" class="grow" id="q" placeholder="Search vendor, invoice no or remarks" value="${esc(st.q)}">
        <select id="fv"><option value="">All vendors</option>${vendorsHere.map((v) => `<option ${v === st.vendor ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select>
        <select id="fs"><option value="due">Sort: due date</option><option value="inv">Sort: newest invoice</option><option value="bal">Sort: highest balance</option><option value="vendor">Sort: vendor</option></select>
      </div>
      <div class="chips" id="chips" style="margin-bottom:12px"></div>
      <div class="totals-bar" id="tot"></div>
      <div class="card"><div class="tbl-wrap" id="list"></div></div>
      <p class="fine">${esc(fyNote())}. Change the financial year at the top to see other years.</p>`;
    $('#fs', el).value = st.sort;
    const draw = () => {
      const q = st.q.toLowerCase();
      const base = rows.filter((r) => (!st.vendor || r.bill.vendorName === st.vendor) &&
        (!q || `${r.bill.vendorName} ${r.bill.invNo} ${r.bill.remarks || ''} ${r.bill.dnNo || ''}`.toLowerCase().includes(q)));
      const counts = Object.fromEntries(STATUS_ORDER.map((s) => [s, 0]));
      base.forEach((r) => counts[r.c.status]++);
      const openN = base.filter((r) => r.c.open).length;
      $('#chips', el).innerHTML = [['ALL', 'All', base.length], ['OPEN', 'All unpaid', openN], ...STATUS_ORDER.map((s) => [s, s.charAt(0) + s.slice(1).toLowerCase(), counts[s]])]
        .map(([k, l, n]) => `<button class="chip ${st.status === k ? 'on' : ''}" data-k="${k}">${esc(l)}<b>${n}</b></button>`).join('');
      $$('#chips .chip', el).forEach((b) => (b.onclick = () => { st.status = b.dataset.k; draw(); }));
      let list = base.filter((r) => st.status === 'ALL' || (st.status === 'OPEN' ? r.c.open : r.c.status === st.status));
      const sorters = {
        due: (a, b) => (a.c.due || '').localeCompare(b.c.due || ''),
        inv: (a, b) => (b.bill.invDate || '').localeCompare(a.bill.invDate || ''),
        bal: (a, b) => b.c.bal - a.c.bal,
        vendor: (a, b) => a.bill.vendorName.localeCompare(b.bill.vendorName) || (a.bill.invDate || '').localeCompare(b.bill.invDate || ''),
      };
      list.sort(sorters[st.sort]);
      const sum = (k) => list.reduce((a, r) => a + r.c[k], 0);
      $('#tot', el).innerHTML = `<span>Bills <b>${list.length}</b></span><span>Invoiced <b>${rs(sum('inv'))}</b></span><span>Paid <b>${rs(sum('paid'))}</b></span><span>${isG ? 'DN' : 'TDS + DN'} <b>${rs(sum('tds') + sum('dn'))}</b></span><span>Balance <b>${rs(sum('bal'))}</b></span>${isG ? `<span>Interest <b>${rs(sum('interest'))}</b></span>` : ''}`;
      $('#list', el).innerHTML = list.length ? `<table class="t cards"><thead><tr><th>Vendor</th><th>Inv no</th><th>Inv date</th><th>Due date</th><th class="num">Inv amt</th><th class="num">Paid</th><th class="num">${isG ? 'DN amt' : 'TDS + DN'}</th><th class="num">Balance</th><th>Days</th><th>Status</th>${isG ? '<th class="num">Interest</th>' : ''}</tr></thead><tbody>
        ${list.map(({ bill: b, c }) => `<tr class="click" data-id="${esc(b.id)}">${td('Vendor', `<span class="vname">${esc(b.vendorName)}</span>`, 'first')}${td('Inv no', esc(b.invNo))}${td('Inv date', dmy(b.invDate))}${td('Due date', dmy(c.due))}${td('Inv amt', rs(c.inv), 'num')}
        ${td('Paid', rs(c.paid) + (c.lastPay ? `<div class="fine">${dmy(c.lastPay)}${(b.payments || []).length > 1 ? ` · ${(b.payments || []).length} instalments` : ''}</div>` : ''), 'num')}${td(isG ? 'DN amt' : 'TDS + DN', rs(c.tds + c.dn), 'num hide-m')}${td('Balance', `<b>${rs(c.bal)}</b>`, 'num')}${td('Days', daysCell(c))}${td('Status', tag(c.status))}${isG ? td('Interest', rs(c.interest), 'num') : ''}</tr>`).join('')}</tbody></table>`
        : '<div class="empty">No bills match these filters.</div>';
      $$('#list tr[data-id]', el).forEach((tr) => (tr.onclick = () => openBill(tr.dataset.id)));
    };
    $('#q', el).oninput = (e) => { st.q = e.target.value; draw(); };
    $('#fv', el).onchange = (e) => { st.vendor = e.target.value; draw(); };
    $('#fs', el).onchange = (e) => { st.sort = e.target.value; draw(); };
    draw();
  };
}
VIEWS.greige = billList(DEPTS.GREIGE);
VIEWS.process = billList(DEPTS.PROCESS);

// ---------------------------------------------------------------- bill detail (read only: the Excel is the master)
function openBill(id) {
  const b = S.bills.find((x) => x.id === id) || (S.allBills || []).find((x) => x.id === id) || (S.ledgerCache._all || []).find((x) => x.id === id);
  if (!b) return toast('Bill not found', 'err');
  const c = computeBill(b, S.settings), isG = b.dept === DEPTS.GREIGE;
  const pays = [...(b.payments || [])].sort((x, y) => (x.date || '').localeCompare(y.date || ''));
  const m = openModal(`
    <div class="detail-head"><div><div class="muted">${DEPT_LABEL[b.dept]} bill</div><h2 style="margin:2px 0">${esc(b.vendorName)}</h2>
      <div>Invoice <b>${esc(b.invNo)}</b> dated <b>${dmy(b.invDate)}</b></div></div>${tag(c.status)}</div>
    <div class="kv">
      <div>Invoice amount<b>${rs2(c.inv)}</b></div><div>Credit terms<b>${esc(b.terms)} days</b></div><div>Due date<b>${dmy(c.due)}</b></div>
      <div>Paid<b>${rs2(c.paid)}</b></div>
      ${isG ? `<div>Debit note${b.dnNo ? ' (No. ' + esc(b.dnNo) + ')' : ''}<b>${rs2(c.dn)}</b></div>` : `<div>TDS + DN<b>${rs2(c.dn + c.tds)}</b></div>`}
      <div>Balance<b class="${c.bal > 0 ? 'neg' : 'pos'}">${rs2(c.bal)}</b></div><div>Timing<b>${esc(daysText(c.days, c.open))}</b></div>
      ${isG ? `<div>Interest on late payments<b>${rs2(c.intPaid)}</b></div><div>Interest accrued on open balance<b>${rs2(c.intAccrued)}</b></div>` : ''}
    </div>
    <div class="subhead"><span>Payments (${pays.length})</span></div>
    ${pays.length ? `<div class="tbl-wrap"><table class="t cards"><thead><tr><th>Date</th><th class="num">Amount</th><th>Mode</th><th>Reference</th><th>Remarks</th><th>Excel row</th></tr></thead><tbody>
      ${pays.map((p) => `<tr>${td('Date', p.date ? dmy(p.date) : '<span class="neg">No date</span>', 'first')}${td('Amount', rs2(p.amount), 'num')}${td('Mode', esc(p.mode || ''))}${td('Reference', esc(p.ref || ''))}${td('Remarks', esc(p.note || ''))}${td('Excel row', p.src === 'excel' ? esc(p.row || '') : '<span class="cat">Recorded in app</span>')}</tr>`).join('')}</tbody></table></div>` : '<p class="muted">No payments yet.</p>'}
    <div class="notice" style="margin-top:14px">This bill comes from the Excel${b.srcRow ? `, <b>${esc(b.srcRow)}</b>` : ''}. To change the bill or its payments, edit the Excel file. The app picks up the change automatically.</div>
    <div class="form-actions"><button class="btn primary" id="closeb">Close</button></div>
    <p class="fine">Last synced ${b.updatedAt ? new Date(b.updatedAt).toLocaleString('en-IN') : ''}</p>`, { wide: true });
  $('#closeb', m.el).onclick = m.close;
}

// ---------------------------------------------------------------- vendors (names and categories come from Excel, details are kept in the app)
VIEWS.vendors = (el) => {
  const st = S.ui.vendors;
  const openBy = new Map();
  for (const r of S.rows) {
    const o = openBy.get(r.bill.vendorId) || { bal: 0, n: 0, urgent: 0 };
    o.bal += r.c.bal; o.n++; if (r.c.status === STATUS.URGENT) o.urgent += r.c.bal;
    openBy.set(r.bill.vendorId, o);
  }
  const counts = { ALL: S.vendors.length };
  for (const k of Object.keys(VCATS)) counts[k] = S.vendors.filter((v) => v.category === k).length;
  el.innerHTML = `
    <div class="toolbar"><input type="search" class="grow" id="vq" placeholder="Search name, city, GSTIN, phone" value="${esc(st.q)}"></div>
    <div class="chips" style="margin-bottom:10px">${[['ALL', 'All'], ...Object.entries(VCATS)].map(([k, l]) => `<button class="chip ${st.cat === k ? 'on' : ''}" data-c="${k}">${l}<b>${counts[k] || 0}</b></button>`).join('')}</div>
    <p class="fine" style="margin:0 0 12px">New vendors are added in the Excel VENDOR MASTER sheet and appear here after the next sync. Contact, GST and bank details are entered here.</p>
    <div class="vgrid" id="vg"></div>`;
  const draw = () => {
    const q = st.q.toLowerCase();
    const list = S.vendors.filter((v) => (st.cat === 'ALL' || v.category === st.cat) &&
      (!q || [v.name, v.city, v.gstin, v.phone, v.contact, v.pan].join(' ').toLowerCase().includes(q)))
      .sort((a, b) => (a.active === false) - (b.active === false) || a.name.localeCompare(b.name));
    $('#vg', el).innerHTML = list.map((v) => {
      const o = openBy.get(v.id);
      return `<div class="vcard c-${esc(v.category)} ${v.active === false ? 'inactive' : ''}" data-id="${esc(v.id)}">
        <div style="display:flex;justify-content:space-between;gap:8px;align-items:flex-start"><span class="nm">${esc(v.name)}</span><span class="cat ${esc(v.category)}">${esc(VCATS[v.category] || v.category)}</span></div>
        <div class="meta">${[v.contact, v.phone, v.city].filter(Boolean).map(esc).join(' · ') || 'No contact details yet'}</div>
        ${v.gstin ? `<div class="meta">GSTIN ${esc(v.gstin)}</div>` : ''}
        ${v.active === false ? `<div class="meta warn-text">Inactive: ${esc(v.inactiveReason || '')}</div>` : ''}
        <div class="bal"><span>${o ? o.n + ' bills' : 'No bills in view'}</span><b class="${o && o.urgent ? 'neg' : ''}">${o ? rs(o.bal) + ' due' : ''}</b></div></div>`;
    }).join('') || '<div class="empty">No vendors found.</div>';
    $$('.vcard', el).forEach((c) => (c.onclick = () => openVendor(c.dataset.id)));
  };
  $('#vq', el).oninput = (e) => { st.q = e.target.value; draw(); };
  $$('.chip', el).forEach((b) => (b.onclick = () => { st.cat = b.dataset.c; render(); }));
  draw();
};
const VFIELDS = [
  ['contact', 'Contact person'], ['phone', 'Phone'], ['email', 'Email'], ['gstin', 'GSTIN'], ['pan', 'PAN'],
  ['address', 'Address'], ['city', 'City'], ['state', 'State'], ['bankName', 'Bank name'], ['accountNo', 'Account no'], ['ifsc', 'IFSC'],
  ['terms', 'Credit terms (days)'], ['notes', 'Notes'],
];
function openVendor(id) {
  const v = vendorById(id); if (!v) return;
  const rows = S.rows.filter((r) => r.bill.vendorId === id);
  const bal = rows.reduce((a, r) => a + r.c.bal, 0);
  const m = openModal(`<div class="detail-head"><div><span class="cat ${esc(v.category)}">${esc(VCATS[v.category] || v.category)}</span><h2 style="margin:6px 0 0">${esc(v.name)}</h2>
     ${v.active === false ? `<span class="tag st-pending">INACTIVE: ${esc(v.inactiveReason || '')}</span>` : ''}</div></div>
    <div class="kv">${VFIELDS.map(([k, l]) => `<div>${l}<b>${esc(v[k] ?? '') || '<span class="muted">Not entered</span>'}</b></div>`).join('')}
      <div>Bills in view<b>${rows.length}</b></div><div>Open balance<b>${rs(bal)}</b></div></div>
    <div class="form-actions"><button class="btn" id="ledv">View ledger</button>
      ${canEdit() ? '<button class="btn primary" id="editv">Edit details</button>' : ''}</div>`, { wide: true });
  $('#ledv', m.el).onclick = () => { m.close(); location.hash = '#/ledger/' + encodeURIComponent(id); };
  if ($('#editv', m.el)) $('#editv', m.el).onclick = () => vendorForm(v);
}
function vendorForm(v) {
  const m = openModal(`<h3>Vendor details</h3><p><b>${esc(v.name)}</b> <span class="cat ${esc(v.category)}">${esc(VCATS[v.category] || v.category)}</span></p>
    <p class="fine">The name and category come from the Excel VENDOR MASTER. Change them there.</p>
    <form id="vf"><div class="form-grid">
    ${VFIELDS.map(([k, l]) => `<label class="f ${k === 'address' || k === 'notes' ? 'span2' : ''}">${l}<input type="${k === 'terms' ? 'number' : k === 'email' ? 'email' : 'text'}" name="${k}" value="${esc(v[k] ?? '')}" ${k === 'gstin' || k === 'pan' || k === 'ifsc' ? 'style="text-transform:uppercase"' : ''}></label>`).join('')}
    </div><p class="warn-text" id="vwarn"></p>
    <div class="form-actions"><button type="button" class="btn ghost" id="vfx">Cancel</button><button class="btn primary">Save details</button></div></form>`, { wide: true });
  const f = $('#vf', m.el);
  $('#vfx', m.el).onclick = () => openVendor(v.id);
  f.onsubmit = async (e) => {
    e.preventDefault();
    const x = formData(f);
    x.gstin = x.gstin.toUpperCase(); x.pan = x.pan.toUpperCase(); x.ifsc = x.ifsc.toUpperCase();
    const warn = [];
    if (x.gstin && !/^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(x.gstin)) warn.push('GSTIN format looks wrong.');
    if (x.pan && !/^[A-Z]{5}\d{4}[A-Z]$/.test(x.pan)) warn.push('PAN format looks wrong.');
    if (x.ifsc && !/^[A-Z]{4}0[A-Z0-9]{6}$/.test(x.ifsc)) warn.push('IFSC format looks wrong.');
    const msg = warn.join(' ') + ' Press save again to keep it.';
    if (warn.length && $('#vwarn', m.el).textContent !== msg) { $('#vwarn', m.el).textContent = msg; return; }
    if (await guard(() => S.store.saveVendor({ id: v.id, ...x }), 'Vendor details saved')) { S.store.addLog('vendor.edit', v.name); setTimeout(() => openVendor(v.id), 120); }
  };
}

// ---------------------------------------------------------------- payment planner
VIEWS.planner = (el) => {
  const st = S.ui.plan;
  el.innerHTML = `
    <div class="card pad" style="margin-bottom:14px"><div class="form-grid">
      <label class="f">Department<select id="pd"><option value="ALL">Both departments</option><option value="GREIGE">Greige</option><option value="PROCESS">Process House</option></select></label>
      <label class="f">Show bills<select id="pw"><option value="over">Overdue on the payment date</option><option value="7">Due within 7 days of payment date</option><option value="15">Due within 15 days</option><option value="30">Due within 30 days</option><option value="all">All unpaid bills</option></select></label>
      <label class="f">Planned payment date<input type="date" id="pdt" value="${st.date}"></label>
      <label class="f">Search<input type="search" id="pq" placeholder="Vendor or invoice" value="${esc(st.q)}"></label>
    </div><p class="fine" style="margin:10px 0 0">Interest shown for greige bills is what would be payable at ${+(S.settings.interestRate * 100).toFixed(2)}% p.a. if the full balance is paid on the planned date.</p></div>
    <div class="card"><div class="tbl-wrap" id="pl"></div></div>
    <div class="sumbar" id="psum"></div>`;
  $('#pd', el).value = st.dept; $('#pw', el).value = st.win;
  let shown = [];
  const draw = () => {
    const pdN = dayNum(st.date || todayStr()), q = st.q.toLowerCase();
    const lim = { over: -1, 7: 7, 15: 15, 30: 30, all: Infinity }[st.win];
    shown = S.rows.filter((r) => r.c.bal > 0 && (st.dept === 'ALL' || r.bill.dept === st.dept) &&
      (dayNum(r.c.due) - pdN <= lim) && (!q || `${r.bill.vendorName} ${r.bill.invNo}`.toLowerCase().includes(q)))
      .map((r) => ({ ...r, int: interestIfPaidOn(r.bill, r.c, st.date, S.settings), dd: dayNum(r.c.due) - pdN }))
      .sort((a, b) => a.c.due.localeCompare(b.c.due));
    const ids = new Set(shown.map((r) => r.bill.id));
    [...st.sel].forEach((id) => { if (!ids.has(id)) st.sel.delete(id); });
    const allOn = shown.length && shown.every((r) => st.sel.has(r.bill.id));
    $('#pl', el).innerHTML = shown.length ? `<table class="t cards" id="ptab"><thead><tr><th><input type="checkbox" class="ck" id="pall" ${allOn ? 'checked' : ''} aria-label="Select all"></th><th>Vendor</th><th>Dept</th><th>Inv no</th><th>Due date</th><th>On payment date</th><th class="num">Balance</th><th class="num">Interest</th><th class="num">Total payable</th></tr></thead><tbody>
      ${shown.map((r) => `<tr class="${st.sel.has(r.bill.id) ? 'sel' : ''}">${td('Select', `<input type="checkbox" class="ck" data-id="${esc(r.bill.id)}" ${st.sel.has(r.bill.id) ? 'checked' : ''}>`)}${td('Vendor', `<span class="vname">${esc(r.bill.vendorName)}</span>`, 'first')}${td('Dept', r.bill.dept === 'GREIGE' ? 'Greige' : 'Process')}${td('Inv no', esc(r.bill.invNo))}${td('Due date', dmy(r.c.due))}
      ${td('On payment date', r.dd < 0 ? `<span class="neg">${-r.dd} days overdue</span>` : r.dd === 0 ? 'Due that day' : `Due ${r.dd} days later`)}${td('Balance', rs(r.c.bal), 'num')}${td('Interest', rs(r.int), 'num')}${td('Total payable', `<b>${rs(r.c.bal + r.int)}</b>`, 'num')}</tr>`).join('')}</tbody></table>`
      : '<div class="empty">No unpaid bills in this window.</div>';
    const sel = shown.filter((r) => st.sel.has(r.bill.id));
    const sb = sel.reduce((a, r) => a + r.c.bal, 0), si = sel.reduce((a, r) => a + r.int, 0);
    const ab = shown.reduce((a, r) => a + r.c.bal, 0);
    $('#psum', el).innerHTML = `<div>Shown<b>${shown.length} bills · ${rs(ab)}</b></div><div>Selected<b>${sel.length} bills</b></div><div>Balance<b>${rs(sb)}</b></div><div>Interest<b>${rs(si)}</b></div><div>Cash needed<b>${rs(sb + si)}</b></div>
      <div class="acts"><button class="btn" id="pvs" ${sel.length ? '' : 'disabled'}>Vendor summary</button><button class="btn" id="pcsv" ${sel.length ? '' : 'disabled'}>Export CSV</button><button class="btn" id="pprint" ${sel.length ? '' : 'disabled'}>Print</button></div>`;
    $$('#pl input[data-id]', el).forEach((c) => (c.onchange = () => { c.checked ? st.sel.add(c.dataset.id) : st.sel.delete(c.dataset.id); draw(); }));
    if ($('#pall', el)) $('#pall', el).onchange = (e) => { shown.forEach((r) => (e.target.checked ? st.sel.add(r.bill.id) : st.sel.delete(r.bill.id))); draw(); };
    const bind = (id, fn) => { const b = $('#' + id, el); if (b) b.onclick = () => fn(sel); };
    bind('pvs', vendorSummary); bind('pcsv', planCSV); bind('pprint', planPrint);
  };
  $('#pd', el).onchange = (e) => { st.dept = e.target.value; draw(); };
  $('#pw', el).onchange = (e) => { st.win = e.target.value; draw(); };
  $('#pdt', el).onchange = (e) => { st.date = e.target.value || todayStr(); draw(); };
  $('#pq', el).oninput = (e) => { st.q = e.target.value; draw(); };
  draw();
};
function groupByVendor(sel) {
  const g = new Map();
  for (const r of sel) {
    const k = r.bill.vendorName; const o = g.get(k) || { name: k, n: 0, bal: 0, int: 0, v: vendorById(r.bill.vendorId) };
    o.n++; o.bal += r.c.bal; o.int += r.int; g.set(k, o);
  }
  return [...g.values()].sort((a, b) => b.bal - a.bal);
}
function vendorSummary(sel) {
  const g = groupByVendor(sel);
  openModal(`<h3>Payment summary by vendor</h3><p class="muted">Planned payment date ${dmy(S.ui.plan.date)}</p>
    <div class="tbl-wrap"><table class="t cards"><thead><tr><th>Vendor</th><th class="num">Bills</th><th class="num">Balance</th><th class="num">Interest</th><th class="num">Total</th><th>Bank details</th></tr></thead><tbody>
    ${g.map((o) => `<tr>${td('Vendor', `<span class="vname">${esc(o.name)}</span>`, 'first')}${td('Bills', o.n, 'num')}${td('Balance', rs(o.bal), 'num')}${td('Interest', rs(o.int), 'num')}${td('Total', `<b>${rs(o.bal + o.int)}</b>`, 'num')}${td('Bank', o.v && o.v.accountNo ? `${esc(o.v.bankName || '')} ${esc(o.v.accountNo)} ${esc(o.v.ifsc || '')}` : '<span class="muted">Not entered</span>')}</tr>`).join('')}</tbody>
    <tfoot><tr>${td('', 'TOTAL', 'first')}${td('Bills', sel.length, 'num')}${td('Balance', rs(g.reduce((a, o) => a + o.bal, 0)), 'num')}${td('Interest', rs(g.reduce((a, o) => a + o.int, 0)), 'num')}${td('Total', rs(g.reduce((a, o) => a + o.bal + o.int, 0)), 'num')}${td('', '')}</tr></tfoot></table></div>`, { wide: true });
}
function planCSV(sel) {
  const rows = [['Planned date', dmy(S.ui.plan.date)], [], ['Vendor', 'Department', 'Invoice no', 'Invoice date', 'Due date', 'Balance', 'Interest', 'Total payable', 'Bank name', 'Account no', 'IFSC']];
  for (const r of sel) {
    const v = vendorById(r.bill.vendorId) || {};
    rows.push([r.bill.vendorName, DEPT_LABEL[r.bill.dept], r.bill.invNo, dmy(r.bill.invDate), dmy(r.c.due), r.c.bal.toFixed(2), r.int.toFixed(2), (r.c.bal + r.int).toFixed(2), v.bankName || '', v.accountNo || '', v.ifsc || '']);
  }
  downloadText(`Payment_plan_${S.ui.plan.date}.csv`, toCSV(rows));
}
function planPrint() { const t = $('#ptab'); t && t.classList.add('print-sel'); window.print(); setTimeout(() => t && t.classList.remove('print-sel'), 500); }
// ---------------------------------------------------------------- vendor ledger
VIEWS.ledger = (el) => {
  const st = S.ui.ledger;
  const withBills = new Set(S.rows.map((r) => r.bill.vendorId));
  const vs = [...S.vendors].sort((a, b) => a.name.localeCompare(b.name));
  const grp = (cat) => vs.filter((v) => v.category === cat).map((v) => `<option value="${esc(v.id)}" ${v.id === st.vendorId ? 'selected' : ''}>${esc(v.name)}${withBills.has(v.id) ? '' : ' (no bills in view)'}</option>`).join('');
  el.innerHTML = `<div class="card pad no-print" style="margin-bottom:14px"><div class="form-grid">
      <label class="f span2">Vendor<select id="lv"><option value="">Select a vendor</option>${Object.entries(VCATS).map(([k, l]) => `<optgroup label="${l}">${grp(k)}</optgroup>`).join('')}</select></label>
      <label class="f">From date<input type="date" id="lf" value="${st.from}"></label>
      <label class="f">To date<input type="date" id="lt" value="${st.to}"></label></div>
      <div class="form-actions" style="justify-content:flex-start"><button class="btn" id="lprint">Print / save as PDF</button><button class="btn" id="lcsv">Export CSV</button></div></div>
    <div id="lbody"><div class="empty">Choose a vendor to see the full account, across all years.</div></div>`;
  $('#lv', el).onchange = (e) => { st.vendorId = e.target.value; history.replaceState(null, '', '#/ledger/' + encodeURIComponent(st.vendorId)); load(); };
  $('#lf', el).onchange = (e) => { st.from = e.target.value; load(); };
  $('#lt', el).onchange = (e) => { st.to = e.target.value; load(); };
  let current = null;
  $('#lprint', el).onclick = () => window.print();
  $('#lcsv', el).onclick = () => {
    if (!current) return toast('Choose a vendor first', 'err');
    downloadText(`Ledger_${current.v.name.replace(/\W+/g, '_')}.csv`, toCSV([[`${current.v.name} ledger`], [], ['Date', 'Particulars', 'Bill amount (Cr)', 'Paid / adjusted (Dr)', 'Balance'],
      ...current.lines.map((l) => [dmy(l.date), l.text, l.cr ? l.cr.toFixed(2) : '', l.dr ? l.dr.toFixed(2) : '', l.bal.toFixed(2)])]));
  };
  const load = async () => {
    const v = vendorById(st.vendorId);
    if (!v) { $('#lbody', el).innerHTML = '<div class="empty">Choose a vendor to see the full account, across all years.</div>'; current = null; return; }
    $('#lbody', el).innerHTML = '<div class="empty">Loading…</div>';
    let bills = S.ledgerCache[v.id];
    if (!bills) { try { bills = await S.store.fetchVendorBills(v.id); } catch { bills = []; } S.ledgerCache[v.id] = bills; }
    const m = new Map(bills.map((b) => [b.id, b]));
    S.bills.filter((b) => b.vendorId === v.id).forEach((b) => m.set(b.id, b)); // live copies win
    const all = [...m.values()];
    S.ledgerCache._all = all;
    const ev = [];
    for (const b of all) {
      ev.push({ date: b.invDate, o: 0, text: `${DEPT_LABEL[b.dept]} bill, inv ${b.invNo}`, cr: +b.invAmt || 0, dr: 0, id: b.id });
      if (+b.dnAmt) ev.push({ date: b.invDate, o: 1, text: `${b.dept === DEPTS.PROCESS ? 'TDS + DN' : 'Debit note' + (b.dnNo ? ' ' + b.dnNo : '')} against inv ${b.invNo}`, cr: 0, dr: +b.dnAmt, id: b.id });
      for (const p of b.payments || []) {
        if (+p.amount) ev.push({ date: p.date, o: 2, text: `Payment${p.mode ? ' by ' + p.mode : ''}${p.ref ? ', ref ' + p.ref : ''} against inv ${b.invNo}`, cr: 0, dr: +p.amount, id: b.id });
        if (+p.tds) ev.push({ date: p.date, o: 3, text: `TDS deducted against inv ${b.invNo}`, cr: 0, dr: +p.tds, id: b.id });
      }
    }
    ev.sort((a, b) => (a.date || '').localeCompare(b.date || '') || a.o - b.o);
    let open = 0; const lines = [];
    for (const e of ev) {
      if (st.from && e.date < st.from) { open += e.cr - e.dr; continue; }
      if (st.to && e.date > st.to) continue;
      lines.push(e);
    }
    let bal = open; lines.forEach((l) => { bal += l.cr - l.dr; l.bal = bal; });
    const cs = all.map((b) => ({ b, c: computeBill(b, S.settings) })).sort((a, b) => a.b.invDate.localeCompare(b.b.invDate));
    const tot = (k) => cs.reduce((a, x) => a + x.c[k], 0);
    current = { v, lines };
    $('#lbody', el).innerHTML = `
      <div class="card pad" style="margin-bottom:14px"><div class="detail-head" style="padding:0"><div><span class="cat ${esc(v.category)}">${esc(VCATS[v.category])}</span><h2 style="margin:6px 0 2px">${esc(v.name)}</h2>
        <div class="muted">${[v.gstin && 'GSTIN ' + v.gstin, v.city, v.phone].filter(Boolean).map(esc).join(' · ')}</div></div>
        <div class="print-only">${esc(S.settings.company)} · Statement ${st.from ? 'from ' + dmy(st.from) : ''} ${st.to ? 'to ' + dmy(st.to) : 'to ' + dmy(todayStr())}</div></div>
        <div class="stat-grid"><div>Total billed<b>${rs(tot('inv'))}</b></div><div>Paid<b>${rs(tot('paid'))}</b></div><div>TDS + DN<b>${rs(tot('tds') + tot('dn'))}</b></div><div>Closing balance<b>${rs(tot('bal'))}</b></div>
        <div>Overdue<b class="neg">${rs(cs.filter((x) => x.c.status === STATUS.URGENT).reduce((a, x) => a + x.c.bal, 0))}</b></div>${v.category === 'GREIGE' || cs.some((x) => x.b.dept === 'GREIGE') ? `<div>Interest (paid late + accrued)<b>${rs(tot('interest'))}</b></div>` : ''}</div></div>
      <div class="card"><div class="sec-title">ACCOUNT STATEMENT</div><div class="tbl-wrap"><table class="t cards"><thead><tr><th>Date</th><th>Particulars</th><th class="num">Bill amount (Cr)</th><th class="num">Paid / adjusted (Dr)</th><th class="num">Balance</th></tr></thead><tbody>
        ${st.from ? `<tr>${td('Date', dmy(st.from), 'first')}${td('Particulars', '<b>Opening balance</b>')}${td('Cr', '', 'num')}${td('Dr', '', 'num')}${td('Balance', rs2(open), 'num')}</tr>` : ''}
        ${lines.map((l) => `<tr class="click" data-id="${esc(l.id)}">${td('Date', dmy(l.date), 'first')}${td('Particulars', esc(l.text))}${td('Bill (Cr)', l.cr ? rs2(l.cr) : '', 'num')}${td('Paid (Dr)', l.dr ? rs2(l.dr) : '', 'num')}${td('Balance', rs2(l.bal), 'num')}</tr>`).join('') || `<tr><td colspan="5" class="empty">No entries.</td></tr>`}</tbody>
        <tfoot><tr>${td('', 'CLOSING BALANCE', 'first')}${td('', '')}${td('Cr', rs2(lines.reduce((a, l) => a + l.cr, 0)), 'num')}${td('Dr', rs2(lines.reduce((a, l) => a + l.dr, 0)), 'num')}${td('Balance', rs2(bal), 'num')}</tr></tfoot></table></div></div>
      <div class="card" style="margin-top:14px"><div class="sec-title">BILL WISE POSITION</div><div class="tbl-wrap"><table class="t cards"><thead><tr><th>Inv no</th><th>Inv date</th><th>Due date</th><th class="num">Amount</th><th class="num">Settled</th><th class="num">Balance</th><th>Status</th></tr></thead><tbody>
        ${cs.map(({ b, c }) => `<tr class="click" data-id="${esc(b.id)}">${td('Inv no', esc(b.invNo), 'first')}${td('Inv date', dmy(b.invDate))}${td('Due date', dmy(c.due))}${td('Amount', rs(c.inv), 'num')}${td('Settled', rs(c.settled), 'num')}${td('Balance', `<b>${rs(c.bal)}</b>`, 'num')}${td('Status', tag(c.status))}</tr>`).join('')}</tbody></table></div></div>`;
    $$('#lbody tr[data-id]', el).forEach((tr) => (tr.onclick = () => openBill(tr.dataset.id)));
  };
  load();
};

// ---------------------------------------------------------------- Excel sync page
VIEWS.data = (el) => {
  const st = od.status(), x = S.sync || {};
  const cfg = od.configured();
  const hrs = st.rtExp ? Math.max(0, Math.round((st.rtExp - Date.now()) / 3600000)) : 0;
  el.innerHTML = `
    <div class="card"><div class="sec-title">1. ONEDRIVE CONNECTION <small>on this device</small></div><div class="pad">
      ${!cfg ? `<div class="notice warn"><b>Microsoft app id not set yet.</b> Follow Part B of the setup guide, then paste the Application (client) ID into <code>js/config.js</code>.
        When registering, use this redirect address:<br><code id="ruri">${esc(od.redirectUri())}</code> <button class="btn sm" id="copyr">Copy</button></div>`
      : st.connected ? `<p>Connected as <b>${esc(st.account?.name || '')}</b> ${esc(st.account?.email || '')}. Microsoft sign in valid for about ${hrs} more hours on this device.</p>`
      : st.expired ? `<div class="notice warn">The OneDrive sign in on this device has expired (Microsoft limits it to 24 hours). Tap Reconnect, it usually completes by itself.</div>`
      : `<p>This device is not connected to OneDrive. Connect it on the device you use most, for example your phone and your laptop.</p>`}
      ${cfg ? `<div class="form-actions" style="justify-content:flex-start">
        ${st.connected ? '<button class="btn" id="odout">Disconnect this device</button>' : `<button class="btn primary" id="odin">${st.expired ? 'Reconnect OneDrive' : 'Connect OneDrive'}</button>`}</div>` : ''}
    </div></div>

    <div class="card"><div class="sec-title">2. EXCEL WORKBOOK</div><div class="pad">
      ${x.itemId ? `<div class="kv"><div>File<b>${esc(x.fileName || '')}</b></div><div>Folder<b>${esc(x.path || '/')}</b></div>
        <div>Last saved in OneDrive<b>${x.fileModified ? new Date(x.fileModified).toLocaleString('en-IN') : ''}</b></div>
        <div>Last synced<b>${x.syncedAt ? new Date(x.syncedAt).toLocaleString('en-IN') : 'Not yet'}</b></div><div>Synced by<b>${esc(x.by || '')}</b></div>
        <div>Checks every<b>${SYNC_CHECK_SECONDS} seconds while open</b></div></div>`
        : '<p class="muted">No workbook chosen yet.</p>'}
      <div class="form-actions" style="justify-content:flex-start">
        ${isAdmin() && st.connected ? `<button class="btn" id="pick">${x.itemId ? 'Change workbook' : 'Choose workbook'}</button>` : ''}
        ${x.itemId && st.connected ? `<button class="btn teal" id="syncnow" ${S.syncBusy ? 'disabled' : ''}>${S.syncBusy ? 'Syncing…' : 'Sync now'}</button>` : ''}
        ${S.pending.length ? `<button class="btn danger" id="review">Review ${S.pending.length} removals</button>` : ''}</div>
      ${S.syncErr ? `<div class="notice warn" style="margin-top:10px">${esc(S.syncErr)}</div>` : ''}
      ${(x.lines || []).length ? `<div class="subhead"><span>Changes in the last sync</span></div><ul class="chg">${x.lines.map((l) => `<li>${esc(l)}</li>`).join('')}</ul>` : ''}
      ${(x.warnings || []).length ? `<div class="notice warn" style="margin-top:10px"><b>${x.warnings.length} points to check in the Excel</b><ul>${x.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul></div>` : ''}
    </div></div>

    <div class="card"><div class="sec-title">3. SYNC FROM A FILE ON THIS DEVICE <small>fallback, same rules as OneDrive</small></div><div class="pad">
      <p class="fine">Use this only if OneDrive is not available. Choose the same workbook.</p>
      <input type="file" id="xf" accept=".xlsx,.xlsm"></div></div>

    <div class="card"><div class="sec-title">4. EXPORT AND BACKUP</div><div class="pad">
      <div class="form-actions" style="justify-content:flex-start"><button class="btn teal" id="xexp">Export to Excel</button>
      ${isAdmin() ? '<button class="btn" id="jexp">Download full backup (JSON)</button>' : ''}
      ${S.store.mode === 'demo' ? '<button class="btn danger" id="dreset">Clear demo data</button>' : ''}</div>
      <p class="fine">The export includes vendor details entered in the app. Keep it private.</p></div></div>`;
  const on = (id, fn) => { const b = $('#' + id, el); if (b) b.onclick = fn; };
  on('copyr', () => { navigator.clipboard?.writeText(od.redirectUri()); toast('Copied'); });
  on('odin', async () => {
    try { const a = await od.connect({ chooseAccount: !st.expired }); if (a !== null) { toast('OneDrive connected'); S.syncErr = ''; render(); chip(); if (S.sync.itemId) runSync({ force: true }); } }
    catch (e) { toast(e.message, 'err'); }
  });
  on('odout', () => { od.disconnect(); toast('Disconnected on this device'); render(); chip(); });
  on('pick', pickWorkbook);
  on('syncnow', () => runSync({ force: true, announce: true }));
  on('review', () => reviewPending());
  on('xexp', async () => guard(async () => exportWorkbook(await S.store.fetchAllBills(), S.vendors, S.settings), 'Export downloaded'));
  on('jexp', async () => guard(async () => {
    const data = { app: 'mkf-payables', exportedAt: new Date().toISOString(), settings: S.settings, vendors: S.vendors, bills: await S.store.fetchAllBills() };
    downloadText(`MKF_backup_${todayStr()}.json`, JSON.stringify(data, null, 1), 'application/json');
  }, 'Backup downloaded'));
  on('dreset', async () => { if (await confirmBox('Clear demo data?', 'All demo data in this browser will be removed.', 'Clear', true)) S.store.resetDemo(); });
  $('#xf', el).onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    runSync({ manual: { buf: await f.arrayBuffer(), name: f.name }, announce: true });
    e.target.value = '';
  };
};

async function pickWorkbook() {
  const m = openModal(`<h3>Choose the workbook in OneDrive</h3>
    <div class="chips" style="margin-bottom:12px"><button class="chip on" data-t="browse">Browse folders</button><button class="chip" data-t="search">Search</button><button class="chip" data-t="link">Paste link</button></div>
    <div id="wp"></div>`, { wide: true });
  const box = $('#wp', m.el);
  const choose = async (f) => {
    if (!(await confirmBox('Use this workbook?', `<b>${esc(f.name)}</b> in ${esc(f.path || '/')}. The app will show exactly what this file contains.`, 'Use it'))) return;
    if (await guard(() => S.store.saveSync({ driveId: f.driveId, itemId: f.itemId, fileName: f.name, path: f.path, cTag: '', lines: [], warnings: [] }), 'Workbook selected')) {
      S.store.addLog('sync.file', f.name);
      S.sync = { ...S.sync, driveId: f.driveId, itemId: f.itemId, fileName: f.name, path: f.path, cTag: '' };
      runSync({ force: true, announce: true });
    }
  };
  const fileRows = (list, emptyMsg) => (list.length ? `<div class="tbl-wrap"><table class="t cards"><tbody>${list.map((f, i) => `<tr class="click" data-i="${i}">
      ${td(f.isFolder ? 'Folder' : 'File', `${f.isFolder ? '📁 ' : '📗 '}<span class="vname">${esc(f.name)}</span>`, 'first')}
      ${td(f.isFolder ? 'Items' : 'Saved', f.isFolder ? `${f.childCount} items` : new Date(f.modified).toLocaleString('en-IN'))}</tr>`).join('')}</tbody></table></div>`
    : `<p class="muted">${emptyMsg}</p>`);
  const bindRows = (list, onFolder) => $$('#wp tr[data-i]', m.el).forEach((tr) => (tr.onclick = () => {
    const f = list[+tr.dataset.i]; f.isFolder ? onFolder(f) : choose(f);
  }));
  const trail = [];
  const browse = async (folder) => {
    box.innerHTML = '<p class="muted">Opening…</p>';
    try {
      const list = await od.listFolder(folder);
      const crumbs = [{ name: 'My OneDrive', f: null }, ...trail.map((t) => ({ name: t.name, f: t }))];
      box.innerHTML = `<div class="fine" style="margin-bottom:8px">${crumbs.map((c, i) => `<a href="javascript:void 0" data-c="${i}">${esc(c.name)}</a>`).join(' › ')}</div>
        ${fileRows(list, 'No folders or Excel files here.')}`;
      $$('a[data-c]', box).forEach((a) => (a.onclick = () => { trail.length = +a.dataset.c; browse(trail[trail.length - 1] || null); }));
      bindRows(list, (f) => { trail.push(f); browse(f); });
    } catch (e) { box.innerHTML = `<div class="notice warn">${esc(e.message)}</div>`; }
  };
  const search = () => {
    box.innerHTML = `<div class="toolbar"><input type="search" class="grow" id="wq" placeholder="Part of the file name" value="Vendor Payments"><button class="btn" id="wgo">Search</button></div><div id="wl"></div>
      <p class="fine">A file uploaded recently may take some time to appear in search. Use Browse folders instead.</p>`;
    const go = async () => {
      $('#wl', box).innerHTML = '<p class="muted">Searching…</p>';
      try {
        const list = await od.findWorkbooks($('#wq', box).value.trim() || 'xlsx');
        $('#wl', box).innerHTML = fileRows(list, 'No Excel files found. Try Browse folders.');
        bindRows(list, () => {});
      } catch (e) { $('#wl', box).innerHTML = `<div class="notice warn">${esc(e.message)}</div>`; }
    };
    $('#wgo', box).onclick = go; $('#wq', box).onkeydown = (e) => { if (e.key === 'Enter') go(); };
    go();
  };
  const link = () => {
    box.innerHTML = `<p class="fine">In OneDrive, select the file, choose <b>Share</b> then <b>Copy link</b> (or open the file and copy the address), and paste it here.</p>
      <div class="toolbar"><input type="text" class="grow" id="wlink" placeholder="https://1drv.ms/x/…  or  https://onedrive.live.com/…"><button class="btn primary" id="wlgo">Open</button></div><div id="wlr"></div>`;
    $('#wlgo', box).onclick = async () => {
      try { choose(await od.fromLink($('#wlink', box).value)); }
      catch (e) { $('#wlr', box).innerHTML = `<div class="notice warn">${esc(e.message)}</div>`; }
    };
  };
  $$('.chip[data-t]', m.el).forEach((c) => (c.onclick = () => {
    $$('.chip[data-t]', m.el).forEach((x) => x.classList.toggle('on', x === c));
    ({ browse: () => browse(trail[trail.length - 1] || null), search, link })[c.dataset.t]();
  }));
  browse(null);
}

// ---------------------------------------------------------------- settings & users
VIEWS.settings = (el) => {
  const s = S.settings;
  const users = [...S.users].sort((a, b) => a.id.localeCompare(b.id));
  el.innerHTML = `
    <div class="card"><div class="sec-title">BUSINESS SETTINGS</div><form class="pad" id="sf"><div class="form-grid">
      <label class="f">Company name<input type="text" name="company" value="${esc(s.company)}"></label>
      <label class="f">Greige late interest (% p.a.)<input type="number" step="0.01" name="interestPct" value="${+(s.interestRate * 100).toFixed(4)}"></label>
      <label class="f">Days in year for interest<input type="number" name="daysInYear" value="${s.daysInYear}"></label>
      <label class="f">Process house grace days<input type="number" name="processGraceDays" value="${s.processGraceDays}"></label>
      <label class="f">Default greige terms (days)<input type="number" name="greigeTerms" value="${s.greigeTerms}"></label>
      <label class="f">Default process house terms (days)<input type="number" name="processTerms" value="${s.processTerms}"></label>
    </div><div class="form-actions"><button class="btn primary">Save settings</button></div></form></div>
    <div class="card"><div class="sec-title">TEAM AND ACCESS <small>Admin: everything · Editor: add and edit bills, payments, vendors · Viewer: read only</small></div><div class="pad">
      <form id="uf" class="form-grid" style="margin-bottom:14px">
        <label class="f">Google email<input type="email" name="email" required placeholder="name@gmail.com"></label>
        <label class="f">Name<input type="text" name="name"></label>
        <label class="f">Role<select name="role"><option value="viewer">Viewer (read only)</option><option value="editor">Editor</option><option value="admin">Admin</option></select></label>
        <div class="f" style="justify-content:flex-end"><button class="btn primary">Add team member</button></div></form>
      <div class="tbl-wrap"><table class="t cards"><thead><tr><th>Email</th><th>Name</th><th>Role</th><th>Status</th><th>PIN</th><th></th></tr></thead><tbody>
      ${users.map((u) => `<tr data-u="${esc(u.id)}">${td('Email', `<span class="vname">${esc(u.id)}</span>`, 'first')}${td('Name', esc(u.name || ''))}
        ${td('Role', `<select data-role ${u.id === S.user.email ? 'disabled' : ''}>${['viewer', 'editor', 'admin'].map((r) => `<option ${u.role === r ? 'selected' : ''}>${r}</option>`).join('')}</select>`)}
        ${td('Status', u.active === false ? '<span class="tag st-pending">DISABLED</span>' : '<span class="tag st-ontime">ACTIVE</span>')}${td('PIN', u.pinHash ? 'Set' : 'Not set')}
        ${td('', u.id === S.user.email ? '<span class="fine">This is you</span>' : `<button class="btn sm" data-act="toggle">${u.active === false ? 'Enable' : 'Disable'}</button> <button class="btn sm" data-act="pin">Reset PIN</button> <button class="btn sm danger" data-act="del">Remove</button>`)}</tr>`).join('')}</tbody></table></div></div></div>
    <div class="card"><div class="sec-title">ACTIVITY LOG <button class="btn sm" id="loadlog" style="color:#fff;background:transparent;border-color:rgba(255,255,255,.4)">Load last 200 actions</button></div><div id="logs" class="tbl-wrap"></div></div>`;
  $('#sf', el).onsubmit = async (e) => {
    e.preventDefault();
    const v = formData(e.target);
    const ns = { company: v.company || 'M K FASHION', interestRate: (Number(v.interestPct) || 0) / 100, daysInYear: Number(v.daysInYear) || 365,
      processGraceDays: Number(v.processGraceDays) || 0, greigeTerms: Number(v.greigeTerms) || 65, processTerms: Number(v.processTerms) || 20 };
    if (await guard(() => S.store.saveSettings(ns), 'Settings saved')) S.store.addLog('settings.save', JSON.stringify(ns));
  };
  $('#uf', el).onsubmit = async (e) => {
    e.preventDefault();
    const v = formData(e.target); const id = v.email.toLowerCase();
    if (S.users.some((u) => u.id === id)) return toast('Already in the team', 'err');
    if (await guard(() => S.store.saveUser({ id, name: v.name, role: v.role, active: true }), 'Team member added')) S.store.addLog('user.add', `${id} as ${v.role}`);
  };
  $$('tr[data-u]', el).forEach((tr) => {
    const id = tr.dataset.u, u = S.users.find((x) => x.id === id);
    const sel = $('select[data-role]', tr);
    if (sel) sel.onchange = async () => { if (await guard(() => S.store.saveUser({ id, role: sel.value }), 'Role updated')) S.store.addLog('user.role', `${id} to ${sel.value}`); };
    $$('button[data-act]', tr).forEach((b) => (b.onclick = async () => {
      const a = b.dataset.act;
      if (a === 'toggle') { if (await guard(() => S.store.saveUser({ id, active: u.active === false }), 'Updated')) S.store.addLog('user.active', `${id} ${u.active === false ? 'enabled' : 'disabled'}`); }
      if (a === 'pin' && (await confirmBox('Reset PIN?', `${esc(id)} will be asked to create a new PIN at next sign in.`, 'Reset'))) { if (await guard(() => S.store.saveUser({ id, pinHash: '' }), 'PIN reset')) S.store.addLog('user.pinreset', id); }
      if (a === 'del' && (await confirmBox('Remove team member?', `${esc(id)} will lose access immediately.`, 'Remove', true))) { if (await guard(() => S.store.deleteUser(id), 'Removed')) S.store.addLog('user.remove', id); }
    }));
  });
  $('#loadlog', el).onclick = async () => {
    const logs = await S.store.fetchLogs().catch((e) => { toast(e.message, 'err'); return []; });
    $('#logs', el).innerHTML = `<table class="t cards"><thead><tr><th>When</th><th>Who</th><th>Action</th><th>Detail</th></tr></thead><tbody>
      ${logs.map((l) => `<tr>${td('When', new Date(l.at).toLocaleString('en-IN'), 'first')}${td('Who', esc(l.by))}${td('Action', esc(l.action))}${td('Detail', esc(l.detail))}</tr>`).join('') || '<tr><td colspan="4" class="empty">No activity yet.</td></tr>'}</tbody></table>`;
  };
};

// ---------------------------------------------------------------- change PIN
function changePin() {
  const m = openModal(`<h3>Change my PIN</h3><form id="cp"><div class="form-grid">
    <label class="f">Current PIN<input type="password" inputmode="numeric" name="old" maxlength="6" required></label>
    <label class="f">New PIN (4 to 6 digits)<input type="password" inputmode="numeric" name="n1" maxlength="6" required></label>
    <label class="f">Repeat new PIN<input type="password" inputmode="numeric" name="n2" maxlength="6" required></label></div><p class="err" id="cperr"></p>
    <div class="form-actions"><button type="button" class="btn ghost" id="cpx">Cancel</button><button class="btn primary">Save PIN</button></div></form>`);
  $('#cpx', m.el).onclick = m.close;
  $('#cp', m.el).onsubmit = async (e) => {
    e.preventDefault();
    const v = formData(e.target), err = $('#cperr', m.el);
    if ((await pinHash(v.old)) !== S.profile.pinHash) return (err.textContent = 'Current PIN is wrong.');
    if (!/^\d{4,6}$/.test(v.n1)) return (err.textContent = 'New PIN must be 4 to 6 digits.');
    if (v.n1 !== v.n2) return (err.textContent = 'New PINs do not match.');
    const h = await pinHash(v.n1);
    if (await guard(() => S.store.setOwnPin(h), 'PIN changed')) { S.profile.pinHash = h; m.close(); }
  };
}

// ============================================================ OneDrive sync engine
let allReady = null;
function ensureAllBills() {
  if (!allReady) {
    allReady = new Promise((res) => {
      let first = true;
      S.store.subscribeAllBills((list) => { S.allBills = list; if (first) { first = false; res(); } });
    });
  }
  return allReady.then(() => S.allBills);
}
let lastErrShown = '';
function startSync() {
  od.completePending().then((a) => { if (a) { toast('OneDrive connected'); if (S.sync.itemId) runSync({ force: true }); } }).catch((e) => toast(e.message, 'err'));
  const tick = () => {
    if (!canEdit() || !S.sync.itemId || document.hidden || !od.configured()) return chip();
    if (!od.status().connected) return chip();
    runSync({});
  };
  setTimeout(tick, 1500);
  setInterval(tick, SYNC_CHECK_SECONDS * 1000);
  setInterval(chip, 30000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) tick(); });
}

async function runSync({ force = false, manual = null, announce = false } = {}) {
  if (S.syncBusy) return;
  S.syncBusy = true; chip();
  try {
    let buf, cTag = S.sync.cTag || '', modified = S.sync.fileModified || '';
    const srcFile = S.sync.itemId || 'manual';
    if (manual) buf = manual.buf;
    else {
      if (!S.sync.itemId) throw new Error('Choose the workbook first.');
      const info = await od.fileInfo(S.sync);
      if (!force && info.cTag && info.cTag === S.sync.cTag) { S.syncErr = ''; return; }   // nothing new saved
      const d = await od.download(S.sync);
      buf = d.buf; cTag = d.cTag || info.cTag; modified = info.modified;
    }
    const parsed = await parseWorkbook(buf);
    if (parsed.exported) throw new Error('This file was exported from the app. Use the working Excel instead.');
    if (!parsed.bills.length) throw new Error('No bills found in this workbook. Check that it is the payables file.');
    const all = await ensureAllBills();
    const plan = planSync(parsed, all, S.vendors, S.settings, { srcFile, isAdmin: isAdmin() });
    if (plan.ops.length) await S.store.writeMany(plan.ops);
    const now = new Date().toISOString();
    const rec = { cTag, fileModified: modified, syncedAt: now, by: S.user.email, warnings: plan.warnings.slice(0, 80), pendingCount: plan.pending.length };
    if (plan.lines.length) Object.assign(rec, { lines: plan.lines.slice(0, 80), more: Math.max(0, plan.lines.length - 80), changedAt: now, counts: plan.counts });
    await S.store.saveSync(rec);
    S.syncErr = ''; lastErrShown = '';
    if (plan.lines.length) S.store.addLog('sync.excel', `${plan.lines.length} changes from ${manual ? manual.name : S.sync.fileName}`);
    else if (announce) toast('Already up to date with the Excel');
    S.pending = plan.pending;
    const tagKey = manual ? 'manual:' + manual.name : cTag;
    let declined = {}; try { declined = JSON.parse(localStorage.getItem('mkf_declined')) || {}; } catch { /* ignore */ }
    const fresh = plan.pending.filter((p) => !(declined.tag === tagKey && (declined.keys || []).includes(p.key)));
    if (fresh.length) reviewPending(tagKey);
  } catch (e) {
    console.error(e);
    S.syncErr = e.reconnect ? e.message : 'Sync problem: ' + e.message;
    if (announce || (!e.reconnect && S.syncErr !== lastErrShown)) { toast(S.syncErr, 'err'); lastErrShown = S.syncErr; }
  } finally {
    S.syncBusy = false; chip();
    if (S.view === 'data') render();
  }
}

function reviewPending(tagKey = S.sync.cTag) {
  const items = S.pending;
  if (!items.length) return;
  const m = openModal(`<h3>Confirm removal</h3>
    <p class="muted">These items are no longer in the Excel. Nothing has been removed yet. Tick what should be removed from the app.</p>
    <div class="pend">${items.map((it, i) => {
      const locked = it.needsAdmin && !isAdmin();
      return `<label class="f chk pend-row"><input type="checkbox" data-i="${i}" ${locked ? 'disabled' : 'checked'}> <span>${esc(it.text)}${locked ? ' <span class="warn-text">(an admin must confirm this)</span>' : ''}</span></label>`;
    }).join('')}</div>
    <div class="form-actions"><button class="btn ghost" id="pnot">Not now</button><button class="btn danger" id="pgo">Remove selected</button></div>`, { wide: true });
  $('#pnot', m.el).onclick = () => {
    localStorage.setItem('mkf_declined', JSON.stringify({ tag: tagKey, keys: items.map((x) => x.key) }));
    m.close(); toast('Kept for now. You will be asked again when the Excel changes.');
  };
  $('#pgo', m.el).onclick = async () => {
    const sel = $$('input[data-i]:checked', m.el).map((c) => items[+c.dataset.i]);
    if (!sel.length) { m.close(); return; }
    const all = await ensureAllBills();
    const ops = removalOps(sel, new Map(all.map((b) => [b.id, b])), S.settings);
    if (await guard(() => S.store.writeMany(ops), `${sel.length} items removed`)) {
      S.store.addLog('sync.remove', sel.map((x) => x.text).join(' | ').slice(0, 480));
      S.pending = items.filter((x) => !sel.includes(x));
      const lines = sel.map((x) => 'Removed: ' + x.text.replace(/ is no longer in the Excel( PAYMENTS sheet)?| was recorded in the app, not in Excel/, ''));
      const now = new Date().toISOString();
      await S.store.saveSync({ lines: [...(S.sync.lines || []), ...lines].slice(-80), changedAt: now, pendingCount: S.pending.length });
      m.close(); if (S.view === 'data') render();
    }
  };
}

const ago = (iso) => {
  if (!iso) return '';
  const s = (Date.now() - new Date(iso).getTime()) / 1000;
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  return new Date(iso).toLocaleDateString('en-IN');
};
function chip() {
  const c = $('#sync-chip'); if (!c) return;
  const x = S.sync || {}, st = od.status();
  let txt = '', cls = '', act = null;
  if (!x.itemId && !x.syncedAt) { if (isAdmin()) { txt = 'Set up Excel sync'; cls = 'warn'; act = () => (location.hash = '#/data'); } }
  else if (S.syncBusy) { txt = 'Syncing…'; cls = 'busy'; }
  else if (canEdit() && od.configured() && st.expired) {
    txt = 'Reconnect OneDrive'; cls = 'warn';
    act = async () => { try { await od.connect(); S.syncErr = ''; runSync({ force: true, announce: true }); } catch (e) { toast(e.message, 'err'); } };
  } else if (S.syncErr && canEdit() && st.connected) { txt = 'Sync problem'; cls = 'warn'; act = () => (location.hash = '#/data'); }
  else { txt = x.syncedAt ? `Synced ${ago(x.syncedAt)}` : 'Waiting for first sync'; cls = st.connected && canEdit() ? 'ok' : ''; act = () => showChanges(); }
  c.hidden = !txt; c.textContent = txt; c.className = 'sync-chip ' + cls; c.onclick = act;
}
function showChanges() {
  const x = S.sync || {};
  openModal(`<h3>Latest changes from Excel</h3><p class="muted">${x.changedAt ? new Date(x.changedAt).toLocaleString('en-IN') : 'No changes recorded yet'}${x.fileName ? ' · ' + esc(x.fileName) : ''}</p>
    ${(x.lines || []).length ? `<ul class="chg">${x.lines.map((l) => `<li>${esc(l)}</li>`).join('')}${x.more ? `<li>and ${x.more} more</li>` : ''}</ul>` : ''}
    <p class="fine">Last checked against the Excel ${ago(x.syncedAt)}.</p>`);
}
function maybeShowSummary() {
  const x = S.sync || {}, card = $('#sync-card');
  if (!x.changedAt || !(x.lines || []).length) return;
  const seen = localStorage.getItem('mkf_seen_change');
  if (seen === x.changedAt) return;
  if (!seen && Date.now() - new Date(x.changedAt).getTime() > 3 * 86400000) { localStorage.setItem('mkf_seen_change', x.changedAt); return; }
  const top = x.lines.slice(0, 6), rest = x.lines.length - top.length + (x.more || 0);
  card.innerHTML = `<div class="sc-head"><b>Updated from Excel</b><span>${ago(x.changedAt)}</span></div>
    <ul class="chg">${top.map((l) => `<li>${esc(l)}</li>`).join('')}</ul>
    <div class="sc-foot">${rest > 0 ? `<button class="btn sm" id="scall">See all ${x.lines.length + (x.more || 0)}</button>` : '<span></span>'}<button class="btn sm primary" id="scok">OK</button></div>`;
  card.hidden = false;
  const done = () => { localStorage.setItem('mkf_seen_change', x.changedAt); card.hidden = true; };
  $('#scok', card).onclick = done;
  if ($('#scall', card)) $('#scall', card).onclick = () => { done(); showChanges(); };
}

boot();
