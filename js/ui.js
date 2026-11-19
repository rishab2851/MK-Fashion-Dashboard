// Small UI helpers: formatting, modal, toast, charts (plain SVG, no library needed).

const inr = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
const inr2 = new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const rs = (v) => (Math.abs(v || 0) < 0.5 ? '₹0' : (v < 0 ? '(₹' + inr.format(-v) + ')' : '₹' + inr.format(v)));
export const rs2 = (v) => '₹' + inr2.format(v || 0);
export const nf = (v) => inr.format(v || 0);
export const pct = (v) => (Number.isFinite(v) ? (v * 100).toFixed(1) + '%' : '0.0%');
export function compact(v) {
  const a = Math.abs(v);
  if (a >= 1e7) return '₹' + (v / 1e7).toFixed(2) + ' Cr';
  if (a >= 1e5) return '₹' + (v / 1e5).toFixed(1) + ' L';
  if (a >= 1e3) return '₹' + (v / 1e3).toFixed(0) + 'K';
  return '₹' + Math.round(v);
}
export const dmy = (s) => (s ? s.split('-').reverse().join('-') : '');
export function daysText(days, open) {
  if (!open) return days >= 0 ? `${days} days before due` : `${-days} days late`;
  if (days < 0) return `${-days} days overdue`;
  if (days === 0) return 'Due today';
  return `Due in ${days} days`;
}

export function $(sel, root = document) { return root.querySelector(sel); }
export function $$(sel, root = document) { return [...root.querySelectorAll(sel)]; }

let toastT;
export function toast(msg, kind = 'ok') {
  const t = $('#toast');
  t.textContent = msg; t.className = 'show ' + kind;
  clearTimeout(toastT); toastT = setTimeout(() => (t.className = ''), 3200);
}

// ---------- modal ----------
export function openModal(html, { wide = false, onClose } = {}) {
  const root = $('#modal-root');
  root.innerHTML = `<div class="modal-back"><div class="modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true">
    <button class="modal-x" aria-label="Close">×</button><div class="modal-body">${html}</div></div></div>`;
  root.hidden = false;
  document.body.classList.add('noscroll');
  const close = () => { root.hidden = true; root.innerHTML = ''; document.body.classList.remove('noscroll'); onClose && onClose(); };
  $('.modal-x', root).onclick = close;
  $('.modal-back', root).addEventListener('mousedown', (e) => { if (e.target.classList.contains('modal-back')) close(); });
  root._close = close;
  return { el: $('.modal-body', root), close };
}
export function closeModal() { const r = $('#modal-root'); r._close && !r.hidden && r._close(); }
export function modalOpen() { return !$('#modal-root').hidden; }

export function confirmBox(title, body, okText = 'Confirm', danger = false) {
  return new Promise((res) => {
    const m = openModal(`<h3>${esc(title)}</h3><div class="confirm-body">${body}</div>
      <div class="form-actions"><button class="btn ghost" data-a="no">Cancel</button>
      <button class="btn ${danger ? 'danger' : 'primary'}" data-a="yes">${esc(okText)}</button></div>`, { onClose: () => res(false) });
    m.el.querySelector('[data-a=no]').onclick = () => { m.close(); };
    m.el.querySelector('[data-a=yes]').onclick = () => { res(true); const r = $('#modal-root'); r._close = null; r.hidden = true; r.innerHTML = ''; document.body.classList.remove('noscroll'); };
  });
}

export function formData(form) {
  const o = {};
  for (const el of form.elements) {
    if (!el.name) continue;
    if (el.type === 'checkbox') o[el.name] = el.checked;
    else if (el.type === 'number') o[el.name] = el.value === '' ? '' : Number(el.value);
    else o[el.name] = el.value.trim();
  }
  return o;
}

export function downloadText(name, text, type = 'text/csv') {
  const blob = new Blob([text], { type: type + ';charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
export function toCSV(rows) {
  return '﻿' + rows.map((r) => r.map((c) => {
    const s = String(c ?? '');
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(',')).join('\n');
}

export async function sha256(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

// ---------- charts ----------
function niceMax(v) {
  if (v <= 0) return 1;
  const p = Math.pow(10, Math.floor(Math.log10(v)));
  const f = v / p;
  return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
}

export function columnChart(items, { color = '#1E6B7B', colors } = {}) {
  const W = 640, Hh = 280, L = 64, R = 12, T = 18, B = 62;
  const max = niceMax(Math.max(...items.map((i) => i.value), 0));
  const cw = (W - L - R) / items.length;
  const bw = Math.min(56, cw * 0.62);
  let g = '';
  for (let i = 0; i <= 4; i++) {
    const y = T + (Hh - T - B) * (1 - i / 4);
    g += `<line x1="${L}" x2="${W - R}" y1="${y}" y2="${y}" class="grid"/><text x="${L - 8}" y="${y + 4}" class="axis" text-anchor="end">${compact(max * i / 4)}</text>`;
  }
  items.forEach((it, i) => {
    const h = (Hh - T - B) * (it.value / max);
    const x = L + cw * i + (cw - bw) / 2;
    const y = Hh - B - h;
    const c = colors ? colors[i] : color;
    g += `<rect x="${x}" y="${y}" width="${bw}" height="${Math.max(h, 0)}" rx="3" fill="${c}"><title>${esc(it.label)}: ${rs(it.value)}</title></rect>`;
    if (it.value > 0) g += `<text x="${x + bw / 2}" y="${y - 5}" class="val" text-anchor="middle">${compact(it.value)}</text>`;
    const words = it.label.split(' ');
    const l1 = words.slice(0, 2).join(' '), l2 = words.slice(2).join(' ');
    g += `<text x="${x + bw / 2}" y="${Hh - B + 18}" class="axis" text-anchor="middle">${esc(l1)}</text>`;
    if (l2) g += `<text x="${x + bw / 2}" y="${Hh - B + 32}" class="axis" text-anchor="middle">${esc(l2)}</text>`;
  });
  return `<svg viewBox="0 0 ${W} ${Hh}" class="chart hbar" role="img">${g}</svg>`;
}

export function stackedBars(rows, series) {
  // rows: [{label, values:[a,b]}], series: [{name,color}]
  if (!rows.length) return '<p class="muted center">No open balances.</p>';
  const W = 760, rowH = 26, L = 200, R = 70, T = 30;
  const Hh = T + rows.length * rowH + 26;
  const max = niceMax(Math.max(...rows.map((r) => r.values.reduce((a, b) => a + b, 0)), 0));
  const sc = (W - L - R) / max;
  let g = series.map((s, i) => `<rect x="${L + i * 150}" y="6" width="12" height="12" rx="2" fill="${s.color}"/><text x="${L + i * 150 + 18}" y="16" class="axis">${esc(s.name)}</text>`).join('');
  for (let i = 0; i <= 4; i++) {
    const x = L + (W - L - R) * i / 4;
    g += `<line x1="${x}" x2="${x}" y1="${T - 4}" y2="${Hh - 22}" class="grid"/><text x="${x}" y="${Hh - 6}" class="axis" text-anchor="middle">${compact(max * i / 4)}</text>`;
  }
  rows.forEach((r, i) => {
    const y = T + i * rowH;
    let x = L;
    const name = r.label.length > 26 ? r.label.slice(0, 25) + '…' : r.label;
    g += `<text x="${L - 8}" y="${y + rowH / 2 + 4}" class="axis" text-anchor="end">${esc(name)}</text>`;
    r.values.forEach((v, j) => {
      const w = v * sc;
      if (w > 0) g += `<rect x="${x}" y="${y + 5}" width="${w}" height="${rowH - 10}" fill="${series[j].color}"><title>${esc(r.label)}, ${esc(series[j].name)}: ${rs(v)}</title></rect>`;
      x += w;
    });
    const tot = r.values.reduce((a, b) => a + b, 0);
    g += `<text x="${x + 6}" y="${y + rowH / 2 + 4}" class="val">${compact(tot)}</text>`;
  });
  return `<svg viewBox="0 0 ${W} ${Hh}" class="chart hbar" role="img">${g}</svg>`;
}
