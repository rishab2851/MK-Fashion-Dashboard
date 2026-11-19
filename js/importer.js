// Excel import (your existing payables workbook) and export.
// Columns are found by their header text, so both the original and the revised workbook import correctly.
import { DEPTS, computeBill, fyOf, billId, vendorIdOf, normName, dayNum, todayStr, fromDayNum } from './calc.js';

const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const EXPORT_MARK = 'EXPORTED FROM M K FASHION PAYABLES APP';

export function loadXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = XLSX_URL; s.onload = () => res(window.XLSX); s.onerror = () => rej(new Error('Could not load the Excel library. Check the internet connection.'));
    document.head.appendChild(s);
  });
}

const H = (v) => String(v ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

function toDate(XLSX, v) {
  if (isNum(v) && v > 20000 && v < 80000) {
    const d = XLSX.SSF.parse_date_code(v);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/); // dd-mm-yyyy
    if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
    const iso = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return iso[0];
  }
  return null;
}

export async function parseWorkbook(buf) {
  const XLSX = await loadXLSX();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false });
  const out = { bills: [], vendors: [], params: {}, warnings: [], sheets: [], exported: false };
  const today = todayStr();

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws['!ref']) continue;
    const rowBase = XLSX.utils.decode_range(ws['!ref']).s.r + 1; // Excel row number of rows[0]
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    if (rows.slice(0, 3).some((r) => r && r.some((c) => String(c ?? '').includes(EXPORT_MARK)))) out.exported = true;
    const up = name.toUpperCase();
    const dept = up.includes('GREIGE') ? DEPTS.GREIGE : up.includes('PROCESS') ? DEPTS.PROCESS : null;

    if (dept) {
      const hi = rows.findIndex((r, i) => i < 15 && r && r.some((c) => H(c) === 'VENDORNAME') && r.some((c) => H(c) === 'INVNO'));
      if (hi < 0) { out.warnings.push(`Sheet "${name}": header row not found, skipped.`); continue; }
      const col = {};
      rows[hi].forEach((c, i) => { const k = H(c); if (k && !(k in col)) col[k] = i; });
      const get = (r, k) => (k in col ? r[col[k]] : null);
      const hasDnAmt = 'DNAMT' in col;
      let n = 0;
      for (let i = hi + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const vendor = normName(get(r, 'VENDORNAME'));
        const invNo = get(r, 'INVNO');
        const invAmt = get(r, 'INVAMT');
        if (!vendor || invNo == null || invNo === '' || !isNum(invAmt)) continue;
        const rowNo = i + rowBase;
        const invDate = toDate(XLSX, get(r, 'INVDATE'));
        if (!invDate) { out.warnings.push(`${name} row ${rowNo}: invoice date missing, bill skipped.`); continue; }
        const terms = isNum(get(r, 'TERMS')) ? get(r, 'TERMS') : (dept === DEPTS.GREIGE ? 65 : 20);
        const paidAmt = isNum(get(r, 'AMTPAID')) ? get(r, 'AMTPAID') : 0;
        let paidOn = toDate(XLSX, get(r, 'PAIDON'));
        let tds = 0, dnAmt = 0, dnNo = '';
        if (dept === DEPTS.PROCESS) {
          const ded = isNum(get(r, 'TDSDN')) ? get(r, 'TDSDN') : 0;
          if (paidAmt > 0) tds = ded; else dnAmt = ded;
          if ('DNNO' in col && get(r, 'DNNO') != null) dnNo = String(get(r, 'DNNO'));
        } else {
          if (hasDnAmt) {
            dnAmt = isNum(get(r, 'DNAMT')) ? get(r, 'DNAMT') : 0;
            dnNo = get(r, 'DNNO') != null ? String(get(r, 'DNNO')) : '';
          } else {
            // original workbook: the "DN NO" column actually held the debit note amount
            dnAmt = isNum(get(r, 'DNNO')) ? get(r, 'DNNO') : 0;
          }
        }
        let payment = null;
        if (paidAmt > 0 || tds > 0) {
          if (!paidOn) { paidOn = today; out.warnings.push(`${name} row ${rowNo} (${vendor} inv ${invNo}): amount paid but no Paid On date. Today's date used, please correct.`); }
          payment = { date: paidOn, amount: paidAmt, tds };
        }
        out.bills.push({ dept, vendorName: vendor, invNo: String(invNo).trim(), invDate, terms, invAmt, dnNo, dnAmt, payment, rowNo, sheet: name });
        n++;
      }
      out.sheets.push(`${name}: ${n} bills`);
      continue;
    }

    if (up.includes('MASTER') || up.includes('SETUP') || up.includes('VENDOR')) {
      const hi = rows.findIndex((r, i) => i < 10 && r && r.some((c) => /SUPPLIER|PROCESS ?HOUSE|TRANSPORT/i.test(String(c ?? ''))));
      if (hi < 0) continue;
      const cats = [];
      rows[hi].forEach((c, i) => {
        const t = String(c ?? '').toUpperCase();
        if (t.includes('SUPPLIER')) cats.push([i, DEPTS.GREIGE]);
        else if (t.includes('PROCESS')) cats.push([i, DEPTS.PROCESS]);
        else if (t.includes('TRANSPORT')) cats.push([i, 'TRANSPORT']);
      });
      let n = 0;
      for (let i = hi + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        for (const [ci, cat] of cats) {
          const v = normName(r[ci]);
          if (v) { out.vendors.push({ name: v, category: cat }); n++; }
        }
        // parameters block (revised workbook)
        r.forEach((c, ci) => {
          const t = String(c ?? '').toLowerCase();
          const val = r[ci + 1];
          if (!isNum(val)) return;
          if (t.includes('interest rate')) out.params.interestRate = val;
          else if (t.includes('days in year')) out.params.daysInYear = val;
          else if (t.includes('grace')) out.params.processGraceDays = val;
        });
      }
      out.sheets.push(`${name}: ${n} vendors`);
    }
  }
  return out;
}

// Work out exactly what an import will change, without writing anything.
export function planImport(parsed, existingBills, existingVendors, settings) {
  const ops = [];
  const sum = { vendorsNew: 0, vendorsFromBills: [], billsNew: 0, billsUpdated: 0, billsSame: 0, conflicts: [], params: parsed.params };
  const vById = new Map(existingVendors.map((v) => [v.id, v]));
  const bById = new Map(existingBills.map((b) => [b.id, b]));

  const wantVendor = (name, category, fromBills) => {
    const id = vendorIdOf(name);
    if (vById.has(id)) return id;
    const v = { id, name, category, active: true, terms: category === DEPTS.GREIGE ? 65 : category === DEPTS.PROCESS ? 20 : '' };
    vById.set(id, v);
    ops.push({ col: 'vendors', id, data: v, merge: true });
    sum.vendorsNew++;
    if (fromBills) sum.vendorsFromBills.push(name);
    return id;
  };
  for (const v of parsed.vendors) wantVendor(v.name, v.category, false);

  const seen = new Set();
  for (const x of parsed.bills) {
    const fy = fyOf(x.invDate);
    const id = billId(x.dept, x.vendorName, fy, x.invNo);
    if (seen.has(id)) { sum.conflicts.push(`${x.sheet} row ${x.rowNo}: same vendor and invoice number as an earlier row (${x.vendorName} inv ${x.invNo}, FY ${fy}). Counted once, this row skipped.`); continue; }
    seen.add(id);
    const vendorId = wantVendor(x.vendorName, x.dept, true);
    const old = bById.get(id);
    const appPays = (old?.payments || []).filter((p) => p.src !== 'excel');
    const xlPay = x.payment ? [{ id: 'xl', date: x.payment.date, amount: x.payment.amount, tds: x.payment.tds, mode: '', ref: '', note: 'Imported from Excel', src: 'excel' }] : [];
    if (appPays.length && xlPay.length) sum.conflicts.push(`${x.vendorName} inv ${x.invNo}: has payments recorded in the app AND in Excel. Both are kept, please check for double counting.`);
    const bill = {
      id, dept: x.dept, vendorId, vendorName: x.vendorName, invNo: x.invNo, invDate: x.invDate, terms: x.terms,
      invAmt: x.invAmt, dnNo: x.dnNo || old?.dnNo || '', dnAmt: x.dnAmt, remarks: old?.remarks || '',
      payments: [...appPays, ...xlPay], fy,
    };
    bill.open = computeBill(bill, settings).open;
    if (!old) sum.billsNew++;
    else {
      const pick = (b) => JSON.stringify([b.vendorName, b.invNo, b.invDate, b.terms, b.invAmt, b.dnNo, b.dnAmt, (b.payments || []).map((p) => [p.date, p.amount, p.tds, p.src])]);
      if (pick(old) === pick(bill)) { sum.billsSame++; continue; }
      sum.billsUpdated++;
    }
    ops.push({ col: 'bills', id, data: bill });
  }
  const changedParams = Object.fromEntries(Object.entries(parsed.params).filter(([k, v]) => settings[k] !== v));
  sum.params = changedParams;
  if (Object.keys(changedParams).length) ops.push({ col: 'settings', id: 'app', data: changedParams, merge: true });
  return { ops, sum };
}

// ---------------- export ----------------
const serial = (s) => (s ? dayNum(s) + 25569 : null);
const dcell = (s) => (s ? { t: 'n', v: serial(s), z: 'dd-mm-yyyy' } : null);
const money = (v) => ({ t: 'n', v: Math.round((v || 0) * 100) / 100, z: '#,##,##0.00' });

export async function exportWorkbook(bills, vendors, settings) {
  const XLSX = await loadXLSX();
  const wb = XLSX.utils.book_new();
  const stampTxt = `${EXPORT_MARK} ON ${fromDayNum(dayNum(todayStr())).split('-').reverse().join('-')}. Do not re-import this file.`;
  const sorted = [...bills].sort((a, b) => (a.invDate || '').localeCompare(b.invDate || '') || a.vendorName.localeCompare(b.vendorName));

  const billSheet = (dept, title) => {
    const isG = dept === DEPTS.GREIGE;
    const head = ['VENDOR NAME', 'INV NO', 'INV DATE', 'TERMS', 'DUE DATE', 'INV AMT', 'LAST PAID ON', 'AMT PAID',
      isG ? 'DN AMT' : 'TDS + DN', 'BALANCE AMT', 'DAYS OVERDUE', 'PAYMENT STATUS', ...(isG ? ['INTEREST (PAID LATE)', 'INTEREST ACCRUED (OPEN)'] : []), 'DN NO', 'REMARKS'];
    const aoa = [[title], [stampTxt], head];
    for (const b of sorted.filter((x) => x.dept === dept)) {
      const c = computeBill(b, settings);
      aoa.push([b.vendorName, b.invNo, dcell(b.invDate), b.terms, dcell(c.due), money(c.inv), dcell(c.lastPay), money(c.paid),
        money(isG ? c.dn : c.tds + c.dn), money(c.bal), c.days, c.status, ...(isG ? [money(c.intPaid), money(c.intAccrued)] : []), b.dnNo || '', b.remarks || '']);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = head.map((h, i) => ({ wch: i === 0 ? 30 : Math.max(12, h.length + 2) }));
    return ws;
  };
  XLSX.utils.book_append_sheet(wb, billSheet(DEPTS.GREIGE, `${settings.company || 'M K FASHION'} GREIGE PAYMENT LIST`), 'GREIGE');
  XLSX.utils.book_append_sheet(wb, billSheet(DEPTS.PROCESS, `${settings.company || 'M K FASHION'} PROCESS HOUSE PAYMENT LIST`), 'PROCESS HOUSE');

  const pay = [['ALL PAYMENTS (INSTALMENT WISE)'], [stampTxt], ['DEPARTMENT', 'VENDOR NAME', 'INV NO', 'PAYMENT DATE', 'AMOUNT', 'TDS / DEDUCTION', 'MODE', 'REFERENCE', 'NOTE']];
  for (const b of sorted) for (const p of b.payments || []) {
    pay.push([b.dept === DEPTS.GREIGE ? 'GREIGE' : 'PROCESS HOUSE', b.vendorName, b.invNo, dcell(p.date), money(p.amount), money(p.tds), p.mode || '', p.ref || '', p.note || '']);
  }
  const wsP = XLSX.utils.aoa_to_sheet(pay); wsP['!cols'] = [16, 30, 10, 14, 14, 16, 12, 18, 30].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsP, 'PAYMENTS');

  const vh = ['NAME', 'CATEGORY', 'CONTACT PERSON', 'PHONE', 'EMAIL', 'GSTIN', 'PAN', 'ADDRESS', 'CITY', 'STATE', 'BANK NAME', 'ACCOUNT NO', 'IFSC', 'CREDIT TERMS', 'ACTIVE', 'NOTES'];
  const va = [['VENDOR MASTER'], [stampTxt], vh];
  for (const v of [...vendors].sort((a, b) => (a.category || '').localeCompare(b.category || '') || a.name.localeCompare(b.name))) {
    va.push([v.name, v.category, v.contact || '', v.phone || '', v.email || '', v.gstin || '', v.pan || '', v.address || '', v.city || '', v.state || '',
      v.bankName || '', v.accountNo || '', v.ifsc || '', v.terms ?? '', v.active === false ? 'NO' : 'YES', v.notes || '']);
  }
  const wsV = XLSX.utils.aoa_to_sheet(va); wsV['!cols'] = vh.map((h, i) => ({ wch: i === 0 ? 30 : 16 }));
  XLSX.utils.book_append_sheet(wb, wsV, 'VENDOR MASTER');

  XLSX.writeFile(wb, `MK_Fashion_Payables_${todayStr()}.xlsx`);
}
