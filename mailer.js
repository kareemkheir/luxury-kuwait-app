import { doc, getDoc, setDoc, getDocs, collection, query, where } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import { calc } from './invoices.js';

// Sends every new invoice to the owner's email as an Excel file, automatically (through a small Google Apps Script the owner sets up once).
let X;
export function initMailer(ctx) { X = ctx; addEventListener('online', () => flushMail()); }
const DAY = 864e5, QK = 'kw_mailq';
const dmy = t => { const d = new Date(t); return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`; };
const loadScript = u => new Promise((res, rej) => { if (window.ExcelJS) return res(); const s = document.createElement('script'); s.src = u; s.onload = res; s.onerror = () => rej(new Error('load')); document.head.appendChild(s); });
const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const getQ = () => { try { return JSON.parse(localStorage.getItem(QK) || '[]'); } catch (e) { return []; } };
const setQ = q => { try { localStorage.setItem(QK, JSON.stringify(q.slice(-30))); } catch (e) {} };
let CFG;
export async function mailConfig(force) {
  if (CFG !== undefined && !force) return CFG;
  try { const d = await getDoc(doc(X.db, 'config', 'mail')); CFG = d.exists() && d.data().url ? d.data() : null; } catch (e) { CFG = null; }
  return CFG;
}

/* how strong is each item for THIS retailer */
async function statuses(inv) {
  const [a, b] = await Promise.all([getDocs(query(collection(X.db, 'invoices'), where('cust', '==', inv.cust))), getDocs(query(collection(X.db, 'returns'), where('cust', '==', inv.cust))).catch(() => ({ docs: [] }))]);
  const prior = a.docs.map(d => d.data()).filter(i => i.no !== inv.no && i.status !== 'cancelled' && i.created < inv.created).sort((x, y) => y.created - x.created);
  const rets = b.docs.map(d => d.data()).filter(r => r.status !== 'cancelled');
  const last3 = prior.slice(0, 3), since = last3.length ? last3[last3.length - 1].created : 0;
  return c => {
    const has = i => i.items.some(x => x.c === c);
    if (rets.some(r => r.created >= since && r.items.some(x => x.c === c))) return ['RETURNED recently', 'r'];
    const n3 = last3.filter(has).length, ls = prior.find(has);
    if (n3 >= 2) return [`STRONG (in ${n3} of his last 3 invoices)`, 'g'];
    if (ls && inv.created - ls.created >= 90 * DAY) return [`SLOW 90+ days (last ordered ${dmy(ls.created)})`, 'o'];
    if (ls) return [`Normal (last ordered ${dmy(ls.created)})`, ''];
    return ['NEW for him', 'b'];
  };
}
const FILL = { r: 'FFF8B4B4', g: 'FFB7E4C0', o: 'FFFFD59A', b: 'FFCFE3FF' };
async function buildMailExcel(inv) {
  await loadScript(EXCELJS);
  let st; try { st = await statuses(inv); } catch (e) { console.error(e); st = () => ['', '']; }
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Invoice');
  ws.columns = [{ width: 24 }, { width: 12 }, { width: 12 }, { width: 46 }];
  ws.mergeCells('A1:D1'); const h = ws.getCell('A1'); h.value = (inv.custEn || '') + (inv.branch ? ' – ' + inv.branch : '') + (inv.custAr ? '   ' + inv.custAr : ''); h.font = { bold: true, size: 16 }; h.alignment = { horizontal: 'center' };
  ws.mergeCells('A2:D2'); const s = ws.getCell('A2'); s.value = `${inv.no} • ${inv.date} • by ${inv.byName || ''} • delivery ${inv.delivery || ''}`; s.alignment = { horizontal: 'center' };
  const hr = ws.addRow([]); const hd = ws.addRow(['Item No.', 'Quantity', 'Price', 'Status for this retailer']); hd.font = { bold: true };
  const first = hd.number + 1;
  inv.items.forEach(it => { const [t, k] = st(it.c); const r = ws.addRow([String(it.c), it.q, it.p, t]); r.getCell(1).numFmt = '@'; r.getCell(3).numFmt = '0.000'; if (FILL[k]) r.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: FILL[k] } }; });
  const last = first + inv.items.length - 1, m = calc(inv.items, inv.dm === 'end' ? inv.disc : 0), qty = inv.items.reduce((a, i) => a + i.q, 0);
  ws.addRow([]);
  const put = (l, v, f) => { const r = ws.addRow([l, v]); r.font = { bold: true }; if (f) r.getCell(2).numFmt = f; };
  put('Final Quantity', qty); put('Subtotal (KWD)', m.sub, '0.000');
  if (inv.dm === 'end' && inv.disc) put(`Discount ${inv.disc}% (KWD)`, -m.da, '0.000');
  put('Final Price (KWD)', m.net, '0.000'); put('Cartons', Math.round(inv.cartons * 100) / 100);
  return await wb.xlsx.writeBuffer();
}
const b64 = buf => { let s = '', u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };

async function post(cfg, p) {
  await fetch(cfg.url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ token: cfg.token, ...p }) });
}
export async function emailInvoice(inv) {
  try {
    const cfg = await mailConfig(); if (!cfg || cfg.on === false) return;
    const buf = await buildMailExcel(inv);
    const p = { subject: `${inv.no} • ${inv.custEn}${inv.branch ? ' – ' + inv.branch : ''} • ${Number(inv.total).toFixed(3)} KWD`, filename: `${inv.custEn}${inv.branch ? ' - ' + inv.branch : ''} ${inv.no}.xlsx`.replace(/[\\/:*?"<>|]/g, ' '), text: `${inv.no} for ${inv.custEn}${inv.branch ? ' (' + inv.branch + ')' : ''}\nTotal ${Number(inv.total).toFixed(3)} KWD\nCreated by ${inv.byName || ''}`, b64: b64(buf) };
    const q = getQ(); q.push({ id: inv.no, p }); setQ(q);
    flushMail();
  } catch (e) { console.error(e); }
}
let busy = false;
export async function flushMail() {
  if (busy || !navigator.onLine) return; const q0 = getQ(); if (!q0.length) return; busy = true;
  try {
    const cfg = await mailConfig(true); if (!cfg || cfg.on === false) return;
    for (const it of q0) { try { await post(cfg, it.p); setQ(getQ().filter(x => x.id !== it.id)); } catch (e) { break; } }
  } finally { busy = false; }
}

/* ---------- one-time setup screen (admin) ---------- */
const script = tok => `const TOKEN = '${tok}';
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  if (d.token !== TOKEN) return ContentService.createTextOutput('no');
  var blob = Utilities.newBlob(Utilities.base64Decode(d.b64), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', d.filename);
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: d.subject, body: d.text, attachments: [blob] });
  return ContentService.createTextOutput('ok');
}`;
export async function mailSetup() {
  const { esc, toast } = X, cfg = await mailConfig(true);
  const tok = (cfg && cfg.token) || Array.from(crypto.getRandomValues(new Uint8Array(18))).map(b => b.toString(16).padStart(2, '0')).join('');
  const m = X.modal(`<h3>Email every invoice to me</h3>
    <div class="sub">${cfg ? '✅ Currently ON. New invoices are emailed to you automatically.' : 'One-time setup (about 3 minutes). Nothing is paid for.'}</div>
    <ol class="sub" style="padding-left:18px;line-height:1.5">
      <li>Open <b>script.google.com</b> signed in with the Gmail that should receive the invoices → <b>New project</b>.</li>
      <li>Delete everything in the editor, tap <b>Copy script</b> below, and paste it. Save.</li>
      <li><b>Deploy → New deployment → ⚙ Web app</b>. Execute as: <b>Me</b>. Who has access: <b>Anyone</b>. Deploy, allow the permissions, and copy the <b>Web app URL</b>.</li>
      <li>Paste that URL here and tap Save.</li></ol>
    <button class="btn l" id="cp">Copy script</button>
    <label>Web app URL</label><input id="ul" placeholder="https://script.google.com/macros/s/…/exec" value="${esc((cfg && cfg.url) || '')}">
    <button class="btn g" id="sv">Save and turn on</button><button class="btn" id="ts">Send a test email</button>
    ${cfg ? '<button class="btn r" id="of">Turn off</button>' : ''}<button class="btn l" id="cl">Close</button>`);
  m.querySelector('#cl').onclick = () => m.remove();
  m.querySelector('#cp').onclick = async () => { try { await navigator.clipboard.writeText(script(tok)); toast('Script copied'); } catch (e) { prompt('Copy this script', script(tok)); } };
  const save = async on => { const url = m.querySelector('#ul').value.trim(); if (!/^https:\/\/script\.google\.com\//.test(url)) return toast('Paste the Web app URL'); await setDoc(doc(X.db, 'config', 'mail'), { url, token: tok, on }); await mailConfig(true); };
  m.querySelector('#sv').onclick = async () => { try { await save(true); toast('Saved. Invoices will be emailed.'); m.remove(); } catch (e) { console.error(e); toast('Could not save'); } };
  m.querySelector('#ts').onclick = async () => {
    try { await save(true); const demo = { no: 'TEST-0001', cust: '-', custEn: 'TEST RETAIL', custAr: '', branch: 'Sample city', date: new Date().toISOString().slice(0, 10), byName: 'Test', delivery: '', items: [{ c: 'ITEM1', q: 2, p: 1.5 }, { c: 'ITEM2', q: 1, p: 4 }], dm: 'end', disc: 10, total: 6.3, cartons: 1 };
      await loadScript(EXCELJS); const buf = await buildMailExcel({ ...demo, cust: '__none__' }); await post(await mailConfig(true), { subject: 'Test invoice email', filename: 'test.xlsx', text: 'If you got this, automatic invoice emails work.', b64: b64(buf) }); toast('Test sent. Check your email in a minute.'); } catch (e) { console.error(e); toast('Could not send the test'); }
  };
  const of = m.querySelector('#of'); if (of) of.onclick = async () => { try { await save(false); toast('Turned off'); m.remove(); } catch (e) { toast('Could not turn off'); } };
}
