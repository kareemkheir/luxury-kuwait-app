import { doc, getDoc, setDoc, getDocs, collection, writeBatch } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import { sendFile } from './invoices.js';

let X;
export function initBackup(ctx) { X = ctx; }
const COLS = ['products', 'customers', 'invoices', 'returns', 'cheques', 'barcodes', 'counters', 'users', 'events', 'config'];
const stamp = () => { const d = new Date(), p = n => String(n).padStart(2, '0'); return `${p(d.getDate())}-${p(d.getMonth() + 1)}-${d.getFullYear()} ${p(d.getHours())}.${p(d.getMinutes())}`; };
const loadScript = u => new Promise((res, rej) => { if (window.ExcelJS) return res(); const s = document.createElement('script'); s.src = u; s.onload = res; s.onerror = () => rej(new Error('load')); document.head.appendChild(s); });
const flat = v => v == null ? '' : (typeof v === 'object' ? JSON.stringify(v) : v);

async function readAll(setMsg) {
  const out = {};
  for (const c of COLS) { setMsg && setMsg(`Reading ${c}…`); try { out[c] = (await getDocs(collection(X.db, c))).docs.map(d => ({ _id: d.id, ...d.data() })); } catch (e) { out[c] = []; } }
  return out;
}
async function makeExcel(data) {
  await loadScript('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js');
  const wb = new ExcelJS.Workbook();
  for (const c of ['customers', 'products', 'invoices', 'returns', 'cheques']) {
    const rows = data[c] || [], keys = [...new Set(rows.flatMap(r => Object.keys(r)))];
    const ws = wb.addWorksheet(c); ws.addRow(keys).font = { bold: true };
    rows.forEach(r => ws.addRow(keys.map(k => flat(r[k])))); ws.views = [{ state: 'frozen', ySplit: 1 }];
  }
  return new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
export async function doBackup(setMsg, mode) {
  const data = await readAll(setMsg); setMsg && setMsg('Building files…');
  const n = stamp();
  await sendFile(new Blob([JSON.stringify({ app: 'luxury-kuwait', at: new Date().toISOString(), data })], { type: 'application/json' }), `Backup ${n}.json`, 'download');
  try { const xl = await makeExcel(data); setTimeout(() => sendFile(xl, `Backup ${n}.xlsx`, 'download'), 900); } catch (e) { console.error(e); }
  try { await setDoc(doc(X.db, 'config', 'lastBackup'), { at: Date.now(), by: X.S.user.name || '' }); } catch (e) {}
  return data;
}

export async function backupHome() {
  const $app = document.getElementById('app'); let last = null;
  try { const d = await getDoc(doc(X.db, 'config', 'lastBackup')); if (d.exists()) last = d.data().at; } catch (e) {}
  const ago = last ? Math.floor((Date.now() - last) / 86400000) : null;
  $app.innerHTML = X.top('Backup • نسخة احتياطية') + `<div class="wrap">
    <div class="card"><b>Last backup:</b> ${last ? new Date(last).toLocaleString('en-GB') + ` (${ago === 0 ? 'today' : ago + ' day' + (ago > 1 ? 's' : '') + ' ago'})` : 'never'}
      <p class="sub">Saves everything (stock, customers, invoices, returns, cheques, barcodes) as 2 files on this phone: an Excel you can open, and a full copy for restoring.</p>
      <button class="btn" id="bk1">💾 Back up now</button><div id="bm" class="sub" style="margin-top:8px"></div></div>
    <div class="card"><b>Best time to back up</b>
      <p class="sub">• Once a week, at the end of the day (after the last invoice).<br>• Always before: deleting items, updating stock from a file, or any big change.<br>• Keep the files in Google Drive or send them to yourself on WhatsApp / email.</p></div>
    <div class="card" style="border:1px solid #c0392b"><b style="color:#c0392b">Delete test data</b>
      <p class="sub">Deletes ALL invoices, returns, cheques and activity history, and restarts invoice numbering from 1. Stock, items, customers (with their opening balances) and users stay. Use once before going live. A backup downloads first.</p>
      <button class="btn" id="bk2" style="background:#c0392b;color:#fff">🗑 Delete test data & reset numbers</button></div></div>`;
  const bm = document.getElementById('bm'), setMsg = t => bm.textContent = t;
  document.getElementById('bk1').onclick = async e => {
    e.target.disabled = true;
    try { await doBackup(setMsg); setMsg('✅ Backup saved to your phone (Downloads).'); } catch (er) { console.error(er); setMsg('❌ Backup failed — try again.'); }
    e.target.disabled = false;
  };
  document.getElementById('bk2').onclick = async e => {
    if (!confirm('Delete ALL invoices, returns, cheques and activity, and reset numbering?\n\nA backup will download first.')) return;
    if ((prompt('Type RESET to confirm') || '').trim() !== 'RESET') { X.toast('Cancelled'); return; }
    e.target.disabled = true;
    try {
      const data = await doBackup(setMsg);
      const refs = [];
      for (const c of ['invoices', 'returns', 'cheques', 'events', 'counters']) (data[c] || []).forEach(r => refs.push(doc(X.db, c, r._id)));
      for (let i = 0; i < refs.length; i += 400) { setMsg(`Deleting… ${Math.min(i + 400, refs.length)}/${refs.length}`); const b = writeBatch(X.db); refs.slice(i, i + 400).forEach(r => b.delete(r)); await b.commit(); }
      setMsg(`✅ Done. ${refs.length} test records deleted; next invoice will be number 1.`);
    } catch (er) { console.error(er); setMsg('❌ Failed — nothing more was changed. Try again.'); }
    e.target.disabled = false;
  };
}
