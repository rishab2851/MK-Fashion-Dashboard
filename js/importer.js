// Reads the payables workbook and works out exactly what changed since the last sync.
// The Excel is the master for bills, payments and vendor names. The app keeps vendor details.
import { DEPTS, computeBill, fyOf, billId, vendorIdOf, normName, dayNum, todayStr } from './calc.js';
import { rs, dmy } from './ui.js';

const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const EXPORT_MARK = 'EXPORTED FROM M K FASHION PAYABLES APP';

export function loadXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = XLSX_URL; s.onload = () => res(window.XLSX); s.onerror = () => rej(new Error('Could not load the Excel reader. Check the internet connection.'));
    document.head.appendChild(s);
  });
}

const H = (v) => String(v ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const r2 = (v) => Math.round((v || 0) * 100) / 100;

function toDate(XLSX, v) {
  if (isNum(v) && v > 20000 && v < 80000) {
    const d = XLSX.SSF.parse_date_code(v);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
    const iso = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return iso[0];
  }
  return null;
}
const invKey = (v) => String(v ?? '').trim().toUpperCase();

function sheetRows(XLSX, ws) {
  if (!ws || !ws['!ref']) return { rows: [], base: 1 };
  return { rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true }), base: XLSX.utils.decode_range(ws['!ref']).s.r + 1 };
}
function headerMap(rows, must, maxRow = 15) {
  const hi = rows.findIndex((r, i) => i < maxRow && r && must.every((k) => r.some((c) => H(c) === k)));
  if (hi < 0) return null;
  const col = {};
  rows[hi].forEach((c, i) => { const k = H(c); if (k && !(k in col)) col[k] = i; });
  return { hi, col };
}

export async function parseWorkbook(buf) {
  const XLSX = await loadXLSX();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false });
  const out = { bills: [], vendors: [], payments: [], params: {}, warnings: [], sheets: [], exported: false, paymentsSheet: false };
  const names = wb.SheetNames;
  const paySheet = names.find((n) => H(n) === 'PAYMENTS');
  out.paymentsSheet = !!paySheet;

  for (const name of names) {
    const { rows, base } = sheetRows(XLSX, wb.Sheets[name]);
    if (rows.slice(0, 3).some((r) => r && r.some((c) => String(c ?? '').includes(EXPORT_MARK)))) out.exported = true;
    const up = name.toUpperCase();
    const dept = up.includes('GREIGE') ? DEPTS.GREIGE : up.includes('PROCESS') ? DEPTS.PROCESS : null;

    if (dept) {
      const hm = headerMap(rows, ['VENDORNAME', 'INVNO']);
      if (!hm) { out.warnings.push(`Sheet "${name}": header row not found, skipped.`); continue; }
      const get = (r, k) => (k in hm.col ? r[hm.col[k]] : null);
      let n = 0;
      for (let i = hm.hi + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const vendor = normName(get(r, 'VENDORNAME'));
        const invNo = get(r, 'INVNO');
        const invAmt = get(r, 'INVAMT');
        if (!vendor || invNo == null || invNo === '' || !isNum(invAmt)) continue;
        const rowNo = i + base;
        const invDate = toDate(XLSX, get(r, 'INVDATE'));
        if (!invDate) { out.warnings.push(`${name} row ${rowNo}: invoice date missing, bill not shown.`); continue; }
        const terms = isNum(get(r, 'TERMS')) ? get(r, 'TERMS') : (dept === DEPTS.GREIGE ? 65 : 20);
        let dnAmt = 0, dnNo = '';
        if (dept === DEPTS.PROCESS) {
          dnAmt = isNum(get(r, 'TDSDN')) ? get(r, 'TDSDN') : 0;              // TDS + DN, deducted at bill level
        } else if ('DNAMT' in hm.col) {
          dnAmt = isNum(get(r, 'DNAMT')) ? get(r, 'DNAMT') : 0;
          dnNo = get(r, 'DNNO') != null ? String(get(r, 'DNNO')) : '';
        } else {
          dnAmt = isNum(get(r, 'DNNO')) ? get(r, 'DNNO') : 0;                 // very first workbook layout
        }
        const bill = { dept, vendorName: vendor, invNo: String(invNo).trim(), invDate, terms, invAmt, dnNo, dnAmt, rowNo, sheet: name, payments: [] };
        if (!paySheet) {                                                      // older workbook: one payment per bill row
          const amt = isNum(get(r, 'AMTPAID')) ? get(r, 'AMTPAID') : 0;
          if (amt > 0) bill.payments.push({ date: toDate(XLSX, get(r, 'PAIDON')), amount: amt, tds: 0, mode: '', ref: '', note: '', row: `${name} row ${rowNo}` });
        }
        out.bills.push(bill); n++;
      }
      out.sheets.push(`${name}: ${n} bills`);
      continue;
    }

    if (name === paySheet) {
      const hm = headerMap(rows, ['DEPARTMENT', 'VENDORNAME', 'INVNO']);
      if (!hm) { out.warnings.push('PAYMENTS sheet: header row not found, payments not read.'); continue; }
      const c = (k) => { const key = Object.keys(hm.col).find((x) => x.startsWith(k)); return key == null ? -1 : hm.col[key]; };
      const ci = { dept: c('DEPARTMENT'), vendor: c('VENDORNAME'), inv: c('INVNO'), date: c('PAYMENTDATE'), amt: c('AMOUNTPAID'), mode: c('MODE'), ref: c('REFERENCE'), note: c('REMARKS') };
      let n = 0;
      for (let i = hm.hi + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const at = (k) => (ci[k] >= 0 ? r[ci[k]] : null);
        const vendor = normName(at('vendor')), inv = at('inv'), amt = at('amt');
        if (!vendor && (inv == null || inv === '') && !isNum(amt)) continue;
        const rowNo = i + base;
        const d = String(at('dept') ?? '').toUpperCase();
        const dept = d.includes('PROCESS') ? DEPTS.PROCESS : d.includes('GREIGE') ? DEPTS.GREIGE : null;
        if (!dept || !vendor || inv == null || inv === '' || !isNum(amt)) { out.warnings.push(`PAYMENTS row ${rowNo}: department, vendor, invoice no or amount missing, row not used.`); continue; }
        const date = toDate(XLSX, at('date'));
        if (!date) out.warnings.push(`PAYMENTS row ${rowNo} (${vendor} inv ${inv}): payment date missing.`);
        out.payments.push({ dept, vendor, inv: String(inv).trim(), date, amount: amt, tds: 0,
          mode: String(at('mode') ?? '').trim(), ref: String(at('ref') ?? '').trim(), note: String(at('note') ?? '').trim(), row: `PAYMENTS row ${rowNo}` });
        n++;
      }
      out.sheets.push(`PAYMENTS: ${n} payments`);
      continue;
    }

    if (up.includes('MASTER') || up === 'SETUP') {
      const hi = rows.findIndex((r, i) => i < 10 && r && r.some((x) => /SUPPLIER|PROCESS ?HOUSE|TRANSPORT/i.test(String(x ?? ''))));
      if (hi < 0) continue;
      const cats = [];
      rows[hi].forEach((x, i) => {
        const t = String(x ?? '').toUpperCase();
        if (t.includes('SUPPLIER')) cats.push([i, DEPTS.GREIGE]);
        else if (t.includes('PROCESS')) cats.push([i, DEPTS.PROCESS]);
        else if (t.includes('TRANSPORT')) cats.push([i, 'TRANSPORT']);
      });
      let n = 0;
      for (let i = hi + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        for (const [ci, cat] of cats) { const v = normName(r[ci]); if (v) { out.vendors.push({ name: v, category: cat }); n++; } }
        r.forEach((x, ci) => {
          const t = String(x ?? '').toLowerCase(); const val = r[ci + 1];
          if (!isNum(val)) return;
          if (t.includes('interest rate')) out.params.interestRate = val;
          else if (t.includes('days in year')) out.params.daysInYear = val;
          else if (t.includes('grace')) out.params.processGraceDays = val;
        });
      }
      out.sheets.push(`${name}: ${n} vendors`);
    }
  }

  // attach PAYMENTS rows to their bills (same department, vendor and invoice number)
  if (paySheet) {
    const idx = new Map();
    for (const b of out.bills) {
      const k = `${b.dept}|${b.vendorName}|${invKey(b.invNo)}`;
      if (!idx.has(k)) idx.set(k, []); idx.get(k).push(b);
    }
    for (const p of out.payments) {
      const list = idx.get(`${p.dept}|${p.vendor}|${invKey(p.inv)}`);
      if (!list) { out.warnings.push(`${p.row}: no ${p.dept === DEPTS.GREIGE ? 'Greige' : 'Process House'} bill found for ${p.vendor} inv ${p.inv}. Payment not counted.`); continue; }
      // same invoice number in two financial years: use the latest bill dated on or before the payment
      const cand = list.filter((b) => !p.date || b.invDate <= p.date).sort((a, b) => b.invDate.localeCompare(a.invDate))[0] || list[0];
      cand.payments.push({ date: p.date, amount: p.amount, tds: 0, mode: p.mode, ref: p.ref, note: p.note, row: p.row });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Sync plan. Nothing is written here. Returns:
//   ops      : safe writes (new bills, changed bills, vendor updates), applied straight away
//   pending  : anything that removes data (a bill, or a payment), applied only after confirmation
//   lines    : plain language summary of every change
const payKey = (p) => `${p.date || ''}|${r2(p.amount)}`;
const payText = (p) => `${rs(p.amount)}${p.date ? ' on ' + dmy(p.date) : ''}`;
const billText = (b) => `${b.vendorName} inv ${b.invNo}`;

export function planSync(parsed, existingBills, vendors, settings, { srcFile, isAdmin }) {
  const ops = [], pending = [], lines = [];
  const counts = { newBills: 0, changedBills: 0, newPayments: 0, newVendors: 0, inactive: 0, reactivated: 0, removals: 0 };
  const vById = new Map(vendors.map((v) => [v.id, v]));
  const bById = new Map(existingBills.map((b) => [b.id, b]));

  // ---------------- vendors
  const master = new Map();
  for (const v of parsed.vendors) { const id = vendorIdOf(v.name); if (!master.has(id)) master.set(id, v); }
  const billVendors = new Map();
  for (const b of parsed.bills) { const id = vendorIdOf(b.vendorName); if (!billVendors.has(id)) billVendors.set(id, b); }
  const touchVendor = (id, name, category, fromMaster) => {
    const v = vById.get(id);
    if (!v) {
      const nv = { id, name, category, active: true, inactiveReason: '', terms: category === DEPTS.GREIGE ? 65 : category === DEPTS.PROCESS ? 20 : '' };
      ops.push({ col: 'vendors', id, data: nv, merge: true }); vById.set(id, nv); counts.newVendors++;
      lines.push(`New vendor: ${name}${fromMaster ? '' : ' (found in bills, not in VENDOR MASTER)'}`);
      return;
    }
    const patch = {};
    if (fromMaster && v.category !== category) patch.category = category;
    if (v.name !== name) patch.name = name;
    if (v.active === false) { patch.active = true; patch.inactiveReason = ''; counts.reactivated++; lines.push(`Vendor active again: ${name}`); }
    if (Object.keys(patch).length) ops.push({ col: 'vendors', id, data: patch, merge: true });
  };
  for (const [id, v] of master) touchVendor(id, v.name, v.category, true);
  for (const [id, b] of billVendors) if (!master.has(id)) touchVendor(id, b.vendorName, b.dept, false);
  if (parsed.vendors.length) {
    for (const v of vendors) {
      if (v.active === false || master.has(v.id) || billVendors.has(v.id)) continue;
      ops.push({ col: 'vendors', id: v.id, data: { active: false, inactiveReason: 'Not in the Excel VENDOR MASTER' }, merge: true });
      counts.inactive++; lines.push(`Vendor marked inactive (removed from Excel): ${v.name}`);
    }
  }

  // ---------------- bills
  const want = new Set();
  for (const x of parsed.bills) {
    const fy = fyOf(x.invDate);
    const id = billId(x.dept, x.vendorName, fy, x.invNo);
    if (want.has(id)) { parsed.warnings.push(`${x.sheet} row ${x.rowNo}: duplicate of an earlier row (${billText(x)}). Shown once.`); continue; }
    want.add(id);
    const old = bById.get(id);
    const xp = [...x.payments].sort((a, b) => (a.date || '').localeCompare(b.date || '') || a.amount - b.amount)
      .map((p, i) => ({ id: 'x' + (i + 1), date: p.date, amount: p.amount, tds: 0, mode: p.mode, ref: p.ref, note: p.note, src: 'excel', row: p.row }));
    const appPays = (old?.payments || []).filter((p) => p.src !== 'excel');
    const bill = {
      id, dept: x.dept, vendorId: vendorIdOf(x.vendorName), vendorName: x.vendorName, invNo: x.invNo, invDate: x.invDate,
      terms: x.terms, invAmt: x.invAmt, dnNo: x.dnNo || '', dnAmt: x.dnAmt, remarks: old?.remarks || '',
      payments: xp, fy, srcFile, srcRow: `${x.sheet} row ${x.rowNo}`,
    };
    bill.open = computeBill(bill, settings).open;

    if (!old) {
      counts.newBills++;
      const paid = xp.reduce((a, p) => a + p.amount, 0);
      lines.push(`New bill: ${billText(bill)}, ${rs(bill.invAmt)}${paid ? `, ${rs(paid)} already paid` : ''}`);
      ops.push({ col: 'bills', id, data: bill });
      continue;
    }
    // what changed on an existing bill
    const ch = [];
    const f = (label, a, b, fmt = (v) => v) => { if (String(a ?? '') !== String(b ?? '')) ch.push(`${label} ${fmt(a) || 'blank'} to ${fmt(b) || 'blank'}`); };
    f('invoice amount', old.invAmt, bill.invAmt, rs); f('invoice date', old.invDate, bill.invDate, dmy); f('terms', old.terms, bill.terms);
    f(x.dept === DEPTS.PROCESS ? 'TDS + DN' : 'DN amount', r2(old.dnAmt), r2(bill.dnAmt), rs); f('DN no', old.dnNo, bill.dnNo);
    const oldX = (old.payments || []).filter((p) => p.src === 'excel');
    const left = [...oldX];
    const added = [];
    for (const p of xp) { const i = left.findIndex((o) => payKey(o) === payKey(p)); if (i >= 0) left.splice(i, 1); else added.push(p); }
    // a payment whose amount changed on the same date is a correction, not a removal
    const corrected = [];
    for (const p of [...added]) {
      const i = left.findIndex((o) => o.date === p.date);
      if (i >= 0) { corrected.push([left[i], p]); left.splice(i, 1); added.splice(added.indexOf(p), 1); }
    }
    added.forEach((p) => lines.push(`Payment ${payText(p)}: ${billText(bill)}`));
    corrected.forEach(([o, p]) => lines.push(`Payment corrected ${payText(o)} to ${rs(p.amount)}: ${billText(bill)}`));
    counts.newPayments += added.length;
    const metaChanged = (old.payments || []).filter((p) => p.src === 'excel').map((p) => [p.mode, p.ref, p.note].join('|')).join(';') !== xp.map((p) => [p.mode, p.ref, p.note].join('|')).join(';');
    const removed = [...left.map((p) => ({ text: `Payment ${payText(p)} on ${billText(bill)} is no longer in the Excel PAYMENTS sheet`, p })),
      ...appPays.map((p) => ({ text: `Payment ${payText(p)} on ${billText(bill)} was recorded in the app, not in Excel`, p }))];
    if (ch.length) lines.push(`${billText(bill)}: ${ch.join(', ')}`);
    const contentChanged = ch.length || added.length || corrected.length || metaChanged;
    if (contentChanged) counts.changedBills++;
    if (removed.length) {
      // write everything except the removals now; the removals wait for confirmation
      const keep = { ...bill, payments: [...xp, ...left, ...appPays] };
      keep.open = computeBill(keep, settings).open;
      ops.push({ col: 'bills', id, data: keep });
      for (const r of removed) pending.push({ key: `pay|${id}|${payKey(r.p)}|${r.p.src}`, text: r.text, billId: id, payment: r.p, needsAdmin: false });
    } else if (contentChanged || old.srcFile !== srcFile || old.srcRow !== bill.srcRow || old.vendorName !== bill.vendorName) {
      ops.push({ col: 'bills', id, data: bill });
    }
  }
  // bills that are no longer in this Excel file
  for (const b of existingBills) {
    if (want.has(b.id)) continue;
    if (b.srcFile && b.srcFile !== srcFile) continue;   // came from another workbook (for example last year's file)
    pending.push({ key: `bill|${b.id}`, text: `Bill ${billText(b)} (${rs(b.invAmt)}, dated ${dmy(b.invDate)}) is no longer in the Excel`, billId: b.id, removeBill: true, needsAdmin: true });
  }
  counts.removals = pending.length;
  // keep the summary readable when a lot arrives at once (for example the very first sync)
  const squash = (prefix, n, text) => {
    const hits = lines.filter((l) => l.startsWith(prefix));
    if (hits.length > n) { const keep = lines.filter((l) => !l.startsWith(prefix)); lines.length = 0; lines.push(text(hits.length), ...keep); }
  };
  squash('New bill:', 15, (n) => `${n} new bills loaded`);
  squash('New vendor:', 10, (n) => `${n} new vendors added to the Vendor Master List`);
  squash('Payment ', 25, (n) => `${n} payments updated`);

  // settings found in the VENDOR MASTER parameters block (admins only)
  if (isAdmin) {
    const p = Object.fromEntries(Object.entries(parsed.params).filter(([k, v]) => settings[k] !== v));
    if (Object.keys(p).length) {
      ops.push({ col: 'settings', id: 'app', data: p, merge: true });
      lines.push(`Settings updated from Excel: ${Object.entries(p).map(([k, v]) => `${k} ${v}`).join(', ')}`);
    }
  }
  return { ops, pending, lines, counts, warnings: parsed.warnings };
}

// Build the writes for the removals the user confirmed.
export function removalOps(selected, billsById, settings) {
  const ops = [];
  const byBill = new Map();
  for (const it of selected) {
    if (it.removeBill) { ops.push({ col: 'bills', id: it.billId, del: true }); continue; }
    if (!byBill.has(it.billId)) byBill.set(it.billId, []);
    byBill.get(it.billId).push(it.payment);
  }
  for (const [id, pays] of byBill) {
    const b = billsById.get(id); if (!b) continue;
    const left = [...(b.payments || [])];
    for (const p of pays) { const i = left.findIndex((o) => o.src === p.src && payKey(o) === payKey(p)); if (i >= 0) left.splice(i, 1); }
    const nb = { ...b, payments: left }; nb.open = computeBill(nb, settings).open;
    ops.push({ col: 'bills', id, data: nb });
  }
  return ops;
}

// ---------------- export (unchanged layout, Process House deductions shown as TDS + DN)
const serial = (s) => (s ? dayNum(s) + 25569 : null);
const dcell = (s) => (s ? { t: 'n', v: serial(s), z: 'dd-mm-yyyy' } : null);
const money = (v) => ({ t: 'n', v: r2(v), z: '#,##,##0.00' });

export async function exportWorkbook(bills, vendors, settings) {
  const XLSX = await loadXLSX();
  const wb = XLSX.utils.book_new();
  const stampTxt = `${EXPORT_MARK} ON ${dmy(todayStr())}. Do not use this file for data entry.`;
  const sorted = [...bills].sort((a, b) => (a.invDate || '').localeCompare(b.invDate || '') || a.vendorName.localeCompare(b.vendorName));
  const billSheet = (dept, title) => {
    const isG = dept === DEPTS.GREIGE;
    const head = ['VENDOR NAME', 'INV NO', 'INV DATE', 'TERMS', 'DUE DATE', 'INV AMT', 'LAST PAID ON', 'AMT PAID',
      isG ? 'DN AMT' : 'TDS + DN', 'BALANCE AMT', 'DAYS OVERDUE', 'PAYMENT STATUS', ...(isG ? ['INTEREST (PAID LATE)', 'INTEREST ACCRUED (OPEN)'] : []), 'DN NO'];
    const aoa = [[title], [stampTxt], head];
    for (const b of sorted.filter((x) => x.dept === dept)) {
      const c = computeBill(b, settings);
      aoa.push([b.vendorName, b.invNo, dcell(b.invDate), b.terms, dcell(c.due), money(c.inv), dcell(c.lastPay), money(c.paid),
        money(c.tds + c.dn), money(c.bal), c.days, c.status, ...(isG ? [money(c.intPaid), money(c.intAccrued)] : []), b.dnNo || '']);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = head.map((h, i) => ({ wch: i === 0 ? 30 : Math.max(12, h.length + 2) }));
    return ws;
  };
  XLSX.utils.book_append_sheet(wb, billSheet(DEPTS.GREIGE, `${settings.company || 'M K FASHION'} GREIGE PAYMENT LIST`), 'GREIGE');
  XLSX.utils.book_append_sheet(wb, billSheet(DEPTS.PROCESS, `${settings.company || 'M K FASHION'} PROCESS HOUSE PAYMENT LIST`), 'PROCESS HOUSE');
  const pay = [['ALL PAYMENTS'], [stampTxt], ['DEPARTMENT', 'VENDOR NAME', 'INV NO', 'PAYMENT DATE', 'AMOUNT', 'MODE', 'REFERENCE', 'REMARKS']];
  for (const b of sorted) for (const p of b.payments || []) pay.push([b.dept === DEPTS.GREIGE ? 'GREIGE' : 'PROCESS HOUSE', b.vendorName, b.invNo, dcell(p.date), money(p.amount), p.mode || '', p.ref || '', p.note || '']);
  const wsP = XLSX.utils.aoa_to_sheet(pay); wsP['!cols'] = [16, 30, 10, 14, 14, 12, 20, 30].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsP, 'PAYMENTS');
  const vh = ['NAME', 'CATEGORY', 'CONTACT PERSON', 'PHONE', 'EMAIL', 'GSTIN', 'PAN', 'ADDRESS', 'CITY', 'STATE', 'BANK NAME', 'ACCOUNT NO', 'IFSC', 'CREDIT TERMS', 'ACTIVE', 'NOTES'];
  const va = [['VENDOR MASTER WITH DETAILS'], [stampTxt], vh];
  for (const v of [...vendors].sort((a, b) => (a.category || '').localeCompare(b.category || '') || a.name.localeCompare(b.name))) {
    va.push([v.name, v.category, v.contact || '', v.phone || '', v.email || '', v.gstin || '', v.pan || '', v.address || '', v.city || '', v.state || '',
      v.bankName || '', v.accountNo || '', v.ifsc || '', v.terms ?? '', v.active === false ? 'NO' : 'YES', v.notes || '']);
  }
  const wsV = XLSX.utils.aoa_to_sheet(va); wsV['!cols'] = vh.map((h, i) => ({ wch: i === 0 ? 30 : 16 }));
  XLSX.utils.book_append_sheet(wb, wsV, 'VENDOR MASTER');
  XLSX.writeFile(wb, `MK_Fashion_Payables_${todayStr()}.xlsx`);
}
