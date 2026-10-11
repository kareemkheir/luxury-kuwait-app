import { getDocs, collection } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import { sendFile } from './invoices.js';

let X;
export function initShelf(ctx) { X = ctx; }
const $ = s => document.querySelector(s);
const DAY = 864e5, SLOW = 60;
const dmy = t => { const d = new Date(t); return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`; };
const loadScript = u => new Promise((res, rej) => { if (window.ExcelJS) return res(); const s = document.createElement('script'); s.src = u; s.onload = res; s.onerror = () => rej(new Error('load')); document.head.appendChild(s); });

let D = null; // { R: Map custId -> retailer model, P: Map code -> product }
async function load(force) {
  if (D && !force) return D;
  const [sv, iv] = await Promise.all([getDocs(collection(X.db, 'shelf')), getDocs(collection(X.db, 'invoices')), X.loadProducts(), X.loadCustomers()]);
  const visits = sv.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => (a.created || 0) - (b.created || 0));
  const invs = iv.docs.map(d => d.data()).filter(i => i.status !== 'cancelled');
  const R = new Map(), P = new Map(X.S.products.map(p => [p.c, p]));
  const get = (id, info) => { if (!R.has(id)) R.set(id, { id, en: (info && info.en) || id, ar: (info && info.ar) || '', visits: [], last: null, first: {}, seen: {}, inv: {} }); return R.get(id); };
  X.S.customers.forEach(c => { if (c.active !== false) get(c.id, c); });
  visits.forEach(v => { const r = get(v.cust, { en: v.custEn, ar: v.custAr }); r.visits.push(v); r.last = v; (v.shelf || []).forEach(c => { r.first[c] ??= v.created; r.seen[c] = v.created; }); });
  invs.forEach(i => { const r = get(i.cust, { en: i.custEn, ar: i.custAr }); i.items.forEach(it => { if ((r.inv[it.c] || 0) < i.created) r.inv[it.c] = i.created; }); });
  const now = Date.now();
  R.forEach(r => { // per item state
    r.items = {}; const lastSet = new Set(r.last ? r.last.shelf || [] : []);
    new Set([...Object.keys(r.seen), ...lastSet]).forEach(c => {
      const base = r.inv[c] || r.first[c], days = Math.floor((now - base) / DAY), now_ = lastSet.has(c);
      r.items[c] = { c, now: now_, lastSeen: r.seen[c], lastInv: r.inv[c] || null, days, slow: now_ && days >= SLOW, never: !r.inv[c] };
    });
    r.have = new Set([...lastSet, ...Object.keys(r.inv).filter(c => now - r.inv[c] < 90 * DAY)]);
  });
  return (D = { R, P, visits });
}
const nm = c => (D.P.get(c) || {}).en || '';
const row = (it, extra = '') => `<div class="card pc">${X.img(it.c)}<div class="pcb"><div class="row"><b>${X.esc(it.c)}</b>${extra}</div><div class="sub">${X.esc(nm(it.c))}</div></div></div>`;

export async function shelfHome() {
  const { top, go, esc } = X;
  $('#app').innerHTML = top('Shelf Report • تقرير الرف') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  let d; try { d = await load(true); } catch (e) { console.error(e); $('#app').innerHTML = top('Shelf Report') + '<div class="wrap"><div class="card">Could not load.</div></div>'; $('#bk').onclick = () => go('home'); return; }
  let tab = 'ret', cust = null, q = '', f = 'now';
  const draw = () => {
    $('#app').innerHTML = top('Shelf Report • تقرير الرف') + `<div class="wrap"><div class="tabs"><button data-t="ret" class="${tab === 'ret' ? 'on' : ''}">Retailer</button><button data-t="item" class="${tab === 'item' ? 'on' : ''}">Item</button><button data-t="sug" class="${tab === 'sug' ? 'on' : ''}">Suggest</button></div><div id="bd"></div>
      <button class="btn l" id="ex">⬇️ Export to Excel</button></div>`;
    $('#bk').onclick = () => go('home');
    document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; cust = null; q = ''; draw(); });
    $('#ex').onclick = exportXl;
    tab === 'ret' ? retTab() : tab === 'item' ? itemTab() : sugTab();
  };
  const custList = (onPick, onlyVisited) => {
    const t = q.trim().toLowerCase();
    return [...d.R.values()].filter(r => (!onlyVisited || r.visits.length) && (!t || r.id.includes(t) || r.en.toLowerCase().includes(t) || (r.ar || '').includes(t))).sort((a, b) => (b.last ? b.last.created : 0) - (a.last ? a.last.created : 0) || a.en.localeCompare(b.en));
  };
  const picker = (note) => {
    $('#bd').innerHTML = `<input id="q" placeholder="Search retailer…" value="${esc(q)}"><div class="sub">${note}</div>` + custList().map(r => {
      const sl = Object.values(r.items).filter(i => i.slow).length;
      return `<div class="card" data-id="${r.id}"><div class="row"><b>${esc(r.en)}</b>${sl ? `<span class="pill o">${sl} slow</span>` : ''}</div><div class="ar sub">${esc(r.ar)}</div><div class="sub">${r.last ? 'Last visit ' + dmy(r.last.created) + ' • ' + (r.last.shelf || []).length + ' on shelf' : 'No shelf check yet'}</div></div>`;
    }).join('');
    const qi = $('#q'); qi.oninput = e => { q = e.target.value; picker(note); const n = $('#q'); n.focus(); n.setSelectionRange(q.length, q.length); };
    document.querySelectorAll('#bd .card[data-id]').forEach(c => c.onclick = () => { cust = c.dataset.id; q = ''; draw(); });
  };
  const retTab = () => {
    if (!cust) return picker('Choose a retailer');
    const r = d.R.get(cust), all = Object.values(r.items);
    const L = { now: all.filter(i => i.now), slow: all.filter(i => i.slow), gone: all.filter(i => !i.now) };
    $('#bd').innerHTML = `<button class="btn l" id="bc" style="margin-top:0">← All retailers</button><div class="card"><b>${esc(r.en)}</b><div class="ar sub">${esc(r.ar)}</div>
      <div class="sub">${r.visits.length} visit(s)${r.last ? ' • last ' + dmy(r.last.created) : ''}</div></div>
      <div class="tabs"><button data-f="now" class="${f === 'now' ? 'on' : ''}">On shelf (${L.now.length})</button><button data-f="slow" class="${f === 'slow' ? 'on' : ''}">Slow (${L.slow.length})</button><button data-f="gone" class="${f === 'gone' ? 'on' : ''}">Not seen last visit (${L.gone.length})</button></div>
      ${L[f].sort((a, b) => b.days - a.days).map(i => row(i, `<span class="pill ${i.slow ? 'o' : 'g'}">${i.slow ? 'slow' : 'ok'}</span>`).replace('</div></div></div>', `</div><div class="sub">${i.lastInv ? 'Last invoiced ' + dmy(i.lastInv) : 'Never invoiced to him'} • last seen ${dmy(i.lastSeen)}</div></div></div>`)).join('') || '<div class="card sub">Nothing here.</div>'}
      <button class="btn" id="gs">💡 What should I sell him?</button>`;
    $('#bc').onclick = () => { cust = null; draw(); };
    document.querySelectorAll('.tabs button[data-f]').forEach(b => b.onclick = () => { f = b.dataset.f; draw(); });
    $('#gs').onclick = () => { tab = 'sug'; draw(); };
  };
  const itemTab = () => {
    const t = q.trim().toLowerCase(), L = t.length < 2 ? [] : X.S.products.filter(p => p.c.toLowerCase().includes(t) || (p.en || '').toLowerCase().includes(t) || (p.ar || '').includes(t)).slice(0, 30);
    if (!cust) {
      $('#bd').innerHTML = `<input id="q" placeholder="Item number or name…" value="${esc(q)}">${L.map(p => `<div class="card pc" data-c="${esc(p.c)}">${X.img(p.c)}<div class="pcb"><b>${esc(p.c)}</b><div class="sub">${esc(p.en)}</div></div></div>`).join('') || '<div class="sub">Type an item to see which retailers have it.</div>'}`;
      const qi = $('#q'); qi.oninput = e => { q = e.target.value; itemTab(); const n = $('#q'); n.focus(); n.setSelectionRange(q.length, q.length); };
      document.querySelectorAll('#bd .card[data-c]').forEach(c => c.onclick = () => { cust = c.dataset.c; draw(); });
      return;
    }
    const has = [], not = [];
    d.R.forEach(r => { if (!r.visits.length) return; const i = r.items[cust]; (i && i.now ? has : not).push({ r, i }); });
    $('#bd').innerHTML = `<button class="btn l" id="bc" style="margin-top:0">← Other item</button>${row({ c: cust })}
      <h3>Has it (${has.length})</h3>${has.map(({ r, i }) => `<div class="card"><div class="row"><b>${esc(r.en)}</b>${i.slow ? '<span class="pill o">slow</span>' : ''}</div><div class="sub">${i.lastInv ? 'Last invoiced ' + dmy(i.lastInv) : 'Never invoiced to him'}</div></div>`).join('') || '<div class="sub">None</div>'}
      <h3>Visited, doesn't have it (${not.length})</h3>${not.map(({ r }) => `<div class="card"><b>${esc(r.en)}</b><div class="sub">last visit ${dmy(r.last.created)}</div></div>`).join('') || '<div class="sub">None</div>'}`;
    $('#bc').onclick = () => { cust = null; q = ''; draw(); };
  };
  const sugTab = () => {
    if (!cust) return picker('Choose a retailer to get suggestions');
    const r = d.R.get(cust), score = new Map(), cnt = new Map(); let n = 0;
    d.R.forEach(o => {
      if (o === r || !o.have.size) return; n++;
      const inter = [...o.have].filter(c => r.have.has(c)).length, w = r.have.size ? (inter + 0.5) / (r.have.size + o.have.size - inter + 0.5) : 0.3;
      o.have.forEach(c => { if (!r.have.has(c)) { score.set(c, (score.get(c) || 0) + w); cnt.set(c, (cnt.get(c) || 0) + 1); } });
    });
    const L = [...score.keys()].filter(c => { const p = d.P.get(c); return p && p.q > 0 && !X.isHidden(p); }).sort((a, b) => score.get(b) - score.get(a)).slice(0, 30);
    $('#bd').innerHTML = `<button class="btn l" id="bc" style="margin-top:0">← All retailers</button><div class="card"><b>${esc(r.en)}</b><div class="sub">Items he doesn't have, that similar shops carry (based on ${n} other retailers). Only items in stock.</div></div>
      ${L.map(c => row({ c }, `<span class="pill g">${cnt.get(c)} shops</span><span class="pill m">stock ${d.P.get(c).q}</span>`)).join('') || '<div class="card sub">Not enough shelf data yet — it improves with every visit.</div>'}`;
    $('#bc').onclick = () => { cust = null; draw(); };
  };
  async function exportXl() {
    try {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js');
      const wb = new ExcelJS.Workbook(), a = wb.addWorksheet('Shelf status'), b = wb.addWorksheet('Visits');
      a.addRow(['Retailer no.', 'Retailer', 'Item', 'Description', 'On shelf at last visit', 'Last seen', 'Last invoiced to him', 'Days (since invoice / first seen)', 'Slow-moving', 'Our stock']).font = { bold: true };
      b.addRow(['Date', 'Retailer', 'By', 'Items on shelf', 'Draft order pieces', 'Invoice']).font = { bold: true };
      d.R.forEach(r => Object.values(r.items).forEach(i => a.addRow([r.id, r.en, i.c, nm(i.c), i.now ? 'Yes' : 'No', dmy(i.lastSeen), i.lastInv ? dmy(i.lastInv) : '', i.days, i.slow ? 'Yes' : 'No', (d.P.get(i.c) || {}).q ?? ''])));
      d.visits.slice().reverse().forEach(v => b.addRow([dmy(v.created), v.custEn, v.byName, (v.shelf || []).length, (v.draft || []).reduce((s, x) => s + x.q, 0), v.invoiced || '']));
      [a, b].forEach(w => { w.views = [{ state: 'frozen', ySplit: 1 }]; w.columns.forEach(c => c.width = 20); });
      await sendFile(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Shelf report ${dmy(Date.now())}.xlsx`, 'download');
    } catch (e) { console.error(e); X.toast('Could not build the Excel file'); }
  }
  draw();
}
