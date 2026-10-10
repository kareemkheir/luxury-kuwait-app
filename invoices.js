import { logEvent } from './activity.js';
import { doc, getDocs, setDoc, collection, runTransaction, query, orderBy, limit } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

let X; // shared context from app.js
export function initInvoices(ctx) { X = ctx; }

const $ = s => document.querySelector(s);
const r3 = n => Math.round(n * 1000) / 1000;
const pad = n => String(n).padStart(5, '0');
const dmy = d => `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const safe = s => String(s).replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
const loaded = {};
function loadScript(url) {
  return loaded[url] ||= new Promise((res, rej) => { const s = document.createElement('script'); s.src = url; s.onload = res; s.onerror = () => rej(new Error('load ' + url)); document.head.appendChild(s); });
}
const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const ZXING = 'https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js';

/* ---------- Excel (shared by invoices and returns) ---------- */
export async function buildExcel(items, name) {
  await loadScript(EXCELJS);
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Invoice');
  ws.columns = [{ width: 24 }, { width: 14 }, { width: 14 }];
  const h = ws.addRow(['Item No.', 'Quantity', 'Price']); h.font = { bold: true };
  items.forEach(it => { const r = ws.addRow([String(it.c), it.q, it.p]); r.getCell(3).numFmt = '0.000'; r.getCell(1).numFmt = '@'; });
  const n = items.length, last = n + 1, t = n + 3;
  const qty = items.reduce((a, i) => a + i.q, 0), price = r3(items.reduce((a, i) => a + i.q * i.p, 0)), cart = items.reduce((a, i) => a + i.q / (i.pk || 1), 0);
  const put = (row, label, formula, result, fmt) => { ws.getCell(`A${row}`).value = label; ws.getCell(`A${row}`).font = { bold: true }; const c = ws.getCell(`B${row}`); c.value = { formula, result }; c.font = { bold: true }; if (fmt) c.numFmt = fmt; };
  put(t, 'Final Quantity', `SUM(B2:B${last})`, qty);
  put(t + 1, 'Final Price (KWD)', `SUMPRODUCT(B2:B${last},C2:C${last})`, price, '0.000');
  put(t + 2, 'Cartons', items.map((it, i) => `B${i + 2}/${it.pk || 1}`).join('+'), cart, '0.##');
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
export async function buildStockExcel(products, inv) {
  await loadScript(EXCELJS);
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Stock');
  ws.columns = [{ width: 18 }, { width: 30 }, { width: 36 }, { width: 10 }, { width: 12 }, { width: 10 }, { width: 22 }, { width: 40 }];
  const h = ws.addRow(['Item Number', 'Arabic Name', 'English Name', 'Packing', 'Stock Qty', 'Price (KWD)', 'Barcode (item number)', 'Added Barcodes']); h.font = { bold: true };
  [...products].sort((a, b) => String(a.c).localeCompare(String(b.c))).forEach(p => {
    const r = ws.addRow([p.c, p.ar || '', p.en || '', p.pk, p.q, p.p, p.c, (inv[p.c] || []).join(', ')]);
    r.getCell(7).numFmt = '@'; r.getCell(8).numFmt = '@'; r.getCell(6).numFmt = '0.000';
  });
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  return new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
export async function sendFile(blob, filename, mode) {
  const f = new File([blob], filename, { type: blob.type });
  if (mode === 'share' && navigator.canShare && navigator.canShare({ files: [f] })) {
    try { await navigator.share({ files: [f], title: filename }); return; } catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ---------- Barcode scanner (camera) ---------- */
export function scanner(onCode, onSnap) {
  const m = X.modal(`<h3>Scan barcode</h3><video id="vd" playsinline muted style="width:100%;border-radius:12px;background:#000;max-height:50vh"></video>
    <div class="sub" id="sm">Starting camera…</div>${onSnap ? '<button class="btn g" id="sn">📸 Read printed item number</button>' : ''}<button class="btn l" id="cl">Close</button>`);
  let reader, stopped = false, last = '', lastT = 0, pend = '', pendN = 0;
  const stop = () => { stopped = true; try { reader && reader.reset(); } catch (e) {} m.remove(); };
  m.querySelector('#cl').onclick = stop;
  const sn = m.querySelector('#sn');
  if (sn) sn.onclick = () => {
    const v = m.querySelector('#vd'); if (!v.videoWidth) return;
    const c = document.createElement('canvas'); c.width = v.videoWidth; c.height = v.videoHeight; c.getContext('2d').drawImage(v, 0, 0);
    stop(); onSnap(c);
  };
  (async () => {
    try {
      await loadScript(ZXING);
      const hints = new Map(), F = ZXing.BarcodeFormat;
      hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E, F.CODE_128, F.CODE_39]);
      hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
      reader = new ZXing.BrowserMultiFormatReader(hints, 120);
      m.querySelector('#sm').textContent = 'Point the camera at the barcode';
      await reader.decodeFromVideoDevice(undefined, m.querySelector('#vd'), res => {
        if (stopped || !res) return;
        const t = res.getText(), now = Date.now();
        // a code counts only when read twice in a row (filters one-frame misreads)
        if (t !== pend) { pend = t; pendN = 1; return; }
        if (++pendN < 2) return;
        if (t === last && now - lastT < 2500) return; last = t; lastT = now;
        const keep = onCode(t); if (keep === false) stop();
      });
    } catch (e) { const s = m.querySelector('#sm'); if (s) s.textContent = 'Camera not available. Allow camera access, or type the item number.'; }
  })();
  return stop;
}


/* ---------- Read a printed item number from a photo (OCR) ---------- */
const TESS = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
let ocrWorker;
async function prepImage(file) {
  const img = await createImageBitmap(file), k = Math.min(2200 / Math.max(img.width, img.height), 2.5);
  const c = document.createElement('canvas'); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
  const x = c.getContext('2d'); x.drawImage(img, 0, 0, c.width, c.height);
  const d = x.getImageData(0, 0, c.width, c.height), g = new Uint8Array(c.width * c.height); let lo = 255, hi = 0;
  for (let i = 0, j = 0; i < d.data.length; i += 4, j++) { const v = (d.data[i] * 0.3 + d.data[i + 1] * 0.59 + d.data[i + 2] * 0.11) | 0; g[j] = v; if (v < lo) lo = v; if (v > hi) hi = v; }
  const r = Math.max(1, hi - lo);
  for (let i = 0, j = 0; i < d.data.length; i += 4, j++) { const v = Math.min(255, Math.max(0, ((g[j] - lo) * 255 / r) | 0)); d.data[i] = d.data[i + 1] = d.data[i + 2] = v; }
  x.putImageData(d, 0, 0); return c;
}
async function ocrText(file) {
  await loadScript(TESS);
  ocrWorker ||= await Tesseract.createWorker('eng');
  await ocrWorker.setParameters({ tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-', tessedit_pageseg_mode: '11' });
  return (await ocrWorker.recognize(await prepImage(file))).data.text;
}
// Compare codes ignoring look-alike letters/digits (O/0, I/L/1, S/5, B/8, Z/2) because OCR confuses them.
const canon = s => String(s).toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1').replace(/S/g, '5').replace(/B/g, '8').replace(/Z/g, '2');
export function matchText(text, products) {
  const idx = new Map(); products.forEach(p => { const k = canon(p.c); if (!idx.has(k)) idx.set(k, []); idx.get(k).push(p); });
  const toks = String(text).split(/[^A-Za-z0-9-]+/).filter(t => t.length >= 3), cands = [];
  toks.forEach((t, i) => { cands.push(t); if (i + 1 < toks.length) cands.push(t + toks[i + 1]); });
  const hits = new Map();
  cands.forEach(t => { const k = canon(t); if (idx.has(k)) idx.get(k).forEach(p => { const exact = p.c.toUpperCase().replace(/[^A-Z0-9]/g, '') === t.toUpperCase().replace(/[^A-Z0-9]/g, ''); const o = hits.get(p.c); if (!o || exact) hits.set(p.c, { p, exact }); }); });
  const all = [...hits.values()], ex = all.filter(h => h.exact);
  return (ex.length ? ex : all).map(h => h.p);
}

/* ---------- New invoice / Return (same screen, two kinds) ---------- */
const KEYS = { invoice: ['cart', 'kw_cart'], return: ['rcart', 'kw_rcart'] };
function saveCart(kind = 'invoice') { const [f, k] = KEYS[kind]; try { localStorage.setItem(k, JSON.stringify(X.S[f] || null)); } catch (e) {} }
function restoreCart(kind) { const [f, k] = KEYS[kind]; if (X.S[f]) return; try { X.S[f] = JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) {} }
const isDamaged = n => { const t = String(n || '').trim().toLowerCase(); return t === 'k' || t === 'ك' || /damag/.test(t) || /تالف/.test(t); };
async function decodeImage(file) {
  await loadScript(ZXING);
  const url = URL.createObjectURL(file);
  try {
    const hints = new Map(); hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
    return (await new ZXing.BrowserMultiFormatReader(hints).decodeFromImageUrl(url)).getText();
  } finally { URL.revokeObjectURL(url); }
}

export async function invoiceHome() { start('invoice'); }
export async function returnHome() { start('return'); }
function start(kind) { restoreCart(kind); const f = KEYS[kind][0]; if (X.S[f] && X.S[f].items) return cartView(kind); customerPick(kind); }

async function customerPick(kind) {
  const { S, top, esc, go, kwd, loadCustomers } = X;
  const R = kind === 'return', F = KEYS[kind][0], TT = R ? 'Return • مرتجع' : 'New Invoice • فاتورة جديدة';
  $('#app').innerHTML = top(TT) + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  const C = await loadCustomers(); let q = '', tab = 'new';
  const draw = () => {
    const t = q.trim().toLowerCase();
    const L = C.filter(c => (c.active !== false || t) && (!t || c.id.includes(t) || (c.en || '').toLowerCase().includes(t) || (c.ar || '').includes(t))).sort((a, b) => (a.en || '').localeCompare(b.en || ''));
    $('#app').innerHTML = top(TT) + `<div class="wrap">
      <div class="tabs"><button data-t="new" class="${tab === 'new' ? 'on' : ''}">New</button><button data-t="recent" class="${tab === 'recent' ? 'on' : ''}">${R ? 'Recent returns' : 'Recent invoices'}</button></div>
      ${tab === 'new' ? `<div class="sub">Choose the customer</div><input id="q" placeholder="Search customer…" value="${esc(q)}">
      ${L.map(c => `<div class="card" data-id="${c.id}"><b>${esc(c.en)}</b><div class="ar sub">${esc(c.ar)}</div>${c.disc ? `<span class="pill g">${c.disc}% discount</span>` : ''}</div>`).join('')}` : '<div id="rl"><div class="sub">Loading…</div></div>'}</div>`;
    $('#bk').onclick = () => go('home');
    document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; draw(); });
    if (tab === 'new') {
      const qi = $('#q'); qi.oninput = e => { q = e.target.value; draw(); const n = $('#q'); n.focus(); n.setSelectionRange(q.length, q.length); };
      document.querySelectorAll('.card[data-id]').forEach(c => c.onclick = () => {
        const cu = C.find(x => x.id === c.dataset.id);
        S[F] = { cust: { id: cu.id, en: cu.en, ar: cu.ar, disc: cu.disc || 0 }, items: [], delivery: iso(new Date(Date.now() + 864e5)), note: '' }; saveCart(kind); cartView(kind);
      });
    } else (R ? recentReturns() : recentList());
  };
  draw();
}

async function cartView(kind) {
  const { S, top, esc, go, kwd, toast, loadProducts } = X;
  const R = kind === 'return', F = KEYS[kind][0], cart = S[F]; await loadProducts();
  await X.loadBarcodes();
  const bkey = c => String(c).trim().replace(/\//g, '_');
  const P = S.products, by = id => { const k = String(id).trim(); let p = P.find(p => p.id === k || p.c.toLowerCase() === k.toLowerCase()); if (!p && S.barcodes[bkey(k)]) p = P.find(p => p.c === S.barcodes[bkey(k)]); return p; };
  const total = () => r3(cart.items.reduce((a, i) => a + i.q * i.p, 0));
  const draw = () => {
    const pcs = cart.items.reduce((a, i) => a + i.q, 0), ctn = cart.items.reduce((a, i) => a + i.q / i.pk, 0);
    $('#app').innerHTML = top((R ? 'Return • ' : 'Invoice • ') + esc(cart.cust.en)) + `<div class="wrap">
      ${cart.cust.disc ? `<div class="pill g" style="margin-bottom:8px">Customer discount ${cart.cust.disc}%</div>` : ''}
      <div class="grid" style="margin-bottom:10px"><button class="btn" id="sc">📷 Scan</button><button class="btn l" id="se">🔎 Search item</button></div>
      <button class="btn l" id="ph" style="margin-top:0">${R ? '🖼 Photo: barcode or printed item number' : '🖼 Read barcode from photo'}</button><input type="file" id="pf" accept="image/*" multiple style="display:none">
      ${cart.items.length === 0 ? '<div class="card sub">No items yet. Scan a barcode or search for an item.</div>' : ''}
      ${cart.items.map((it, i) => `<div class="card"><div class="row"><b>${esc(it.c)}</b><button class="pill r" data-rm="${i}" style="border:0">Remove</button></div>
        <div class="sub">${esc(it.en)}</div><div class="ar sub">${esc(it.ar)}</div>
        <div class="row" style="margin-top:6px"><div style="flex:1"><label>${R ? 'Pieces returned' : 'Pieces (stock ' + it.stock + ')'}</label><input type="number" inputmode="numeric" data-q="${i}" value="${it.q}"></div>
        <button class="btn l" style="width:auto;margin:20px 0 0" data-ctn="${i}">+1 carton<br><span class="sub">(${it.pk} pcs)</span></button></div>
        ${R ? `<div class="row"><input data-nt="${i}" placeholder="Note — write damaged / تالف / K if broken" value="${esc(it.note || '')}" style="flex:1"><button class="btn l" style="width:auto;margin:6px 0" data-dm="${i}">🔨</button></div>${isDamaged(it.note) ? '<span class="pill r">Damaged — will NOT go back to stock</span>' : '<span class="pill g">Goes back to stock</span>'}` : ''}
        <div class="sub">${kwd(it.p)} KWD each • ${(it.q / it.pk).toFixed(2)} cartons • <b>${kwd(it.q * it.p)}</b></div></div>`).join('')}
      <div class="card"><div class="row"><span>Pieces</span><b>${pcs}</b></div><div class="row"><span>Cartons</span><b>${ctn.toFixed(2)}</b></div>
        <div class="row"><span>Total (KWD)</span><b style="font-size:22px">${kwd(total())}</b></div></div>
      ${R ? '' : `<label>Delivery date • تاريخ التوصيل</label><input type="date" id="dd" value="${cart.delivery}">`}
      <label>Note (optional)</label><input id="nt" value="${esc(cart.note)}">
      <button class="btn g" id="cr" ${cart.items.length ? '' : 'disabled style="opacity:.5"'}>${R ? 'Create return' : 'Create invoice'}</button>
      <button class="btn l" id="cn">Cancel / start over</button></div>`;
    $('#bk').onclick = () => { go('home'); };
    document.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { cart.items.splice(+b.dataset.rm, 1); saveCart(kind); draw(); });
    document.querySelectorAll('[data-q]').forEach(inp => inp.onchange = () => {
      const it = cart.items[+inp.dataset.q]; let v = Math.max(1, Math.floor(+inp.value || 1));
      if (!R && v > it.stock) { v = it.stock; toast(`Only ${it.stock} in stock`); } it.q = v; saveCart(kind); draw();
    });
    document.querySelectorAll('[data-ctn]').forEach(b => b.onclick = () => {
      const it = cart.items[+b.dataset.ctn]; if (!R && it.q + it.pk > it.stock) { toast(`Only ${it.stock} in stock`); it.q = it.stock; } else it.q += it.pk; saveCart(kind); draw();
    });
    const dd = $('#dd'); if (dd) dd.onchange = e => { cart.delivery = e.target.value; saveCart(kind); };
    document.querySelectorAll('[data-nt]').forEach(inp => inp.onchange = () => { cart.items[+inp.dataset.nt].note = inp.value; saveCart(kind); draw(); });
    document.querySelectorAll('[data-dm]').forEach(b => b.onclick = () => { cart.items[+b.dataset.dm].note = 'damaged'; saveCart(kind); draw(); });
    $('#nt').onchange = e => { cart.note = e.target.value; saveCart(kind); };
    $('#cn').onclick = () => { if (!cart.items.length || confirm('Throw away this draft?')) { S[F] = null; saveCart(kind); go(R ? 'return' : 'invoice'); } };
    $('#sc').onclick = () => scanner(code => {
      if (!by(code)) { setTimeout(() => linkBarcode(code), 50); return false; }
      toast(addItem(code)); return true;
    }, snapRead);
    $('#se').onclick = searchItem;
    async function snapRead(cv) {
      try {
        toast('Reading the printed number…');
        const hits = matchText(await ocrText(cv), S.products);
        if (hits.length === 1) return toast(addItem(hits[0].c));
        if (hits.length > 1) return pickOne(hits, null);
        toast('Could not read an item number. Get closer so the number fills the screen, or use Search.');
      } catch (err) { console.error(err); toast('Could not read the photo'); }
    }
    $('#cr').onclick = R ? createReturn : create;
    $('#ph').onclick = () => $('#pf').click();
    $('#pf').onchange = async e => {
      const fs = [...e.target.files]; e.target.value = ''; if (!fs.length) return;
      let ok = 0, bad = 0, unknown = null;
      for (const f of fs) {
        toast('Reading…'); let code = null;
        try { code = await decodeImage(f); } catch (err) {}
        if (code && by(code)) { const r = addItem(code); if (r.startsWith('Added')) ok++; else toast(r); continue; }
        {   // also try to read the printed item number
          try {
            toast('Reading the printed number… (first time takes a few seconds)');
            const hits = matchText(await ocrText(f), S.products);
            if (hits.length === 1) {
              const p = hits[0]; addItem(p.c); ok++;
              if (code) { try { await setDoc(doc(X.db, 'barcodes', bkey(code)), { c: p.c, code: String(code).trim(), by: S.user.name, at: Date.now() }); S.barcodes[bkey(code)] = p.c; (S.barcodeCodes ||= {})[bkey(code)] = String(code).trim(); } catch (e2) {} }
              continue;
            }
            if (hits.length > 1) { pickOne(hits, code); continue; }
          } catch (err) { console.error(err); }
        }
        if (code) unknown ||= code; else bad++;
      }
      toast(`${ok} item(s) added${bad ? ` • nothing readable in ${bad} photo(s) — try a closer, sharper photo, or use Search` : ''}`);
      if (unknown) linkBarcode(unknown);
    };
  };
  function pickOne(list, code) {
    const m = X.modal(`<h3>Which item is it?</h3><div class="sub">The photo matched more than one item. Tap the right one.</div>${list.map(p => `<div class="card" data-c="${esc(p.c)}"><b>${esc(p.c)}</b><div class="sub">${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div></div>`).join('')}<button class="btn l" id="cl">Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelectorAll('[data-c]').forEach(c => c.onclick = async () => {
      addItem(c.dataset.c); m.remove();
      if (code) { try { await setDoc(doc(X.db, 'barcodes', bkey(code)), { c: c.dataset.c, code: String(code).trim(), by: S.user.name, at: Date.now() }); S.barcodes[bkey(code)] = c.dataset.c; (S.barcodeCodes ||= {})[bkey(code)] = String(code).trim(); } catch (e2) {} }
    });
  }
  function linkBarcode(code) {
    const m = X.modal(`<h3>Barcode not found</h3><div class="card"><div class="sub">The camera read this number:</div><b style="font-size:20px;word-break:break-all">${esc(code)}</b></div>
      <div class="sub">Search for the item once and tap it. Next time this barcode adds it automatically.</div>
      <input id="si" placeholder="Item number or name…"><div id="sr"></div><button class="btn l" id="cl">Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelector('#si').oninput = e => {
      const t = e.target.value.trim().toLowerCase(); if (t.length < 2) { m.querySelector('#sr').innerHTML = ''; return; }
      const L = P.filter(p => p.c.toLowerCase().includes(t) || (p.en || '').toLowerCase().includes(t) || (p.ar || '').includes(t)).slice(0, 25);
      m.querySelector('#sr').innerHTML = L.map(p => `<div class="card" data-c="${esc(p.c)}"><div class="row"><b>${esc(p.c)}</b><span class="pill ${p.q <= 0 ? 'r' : 'g'}">${p.q <= 0 ? 'OUT' : p.q}</span></div><div class="sub">${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div></div>`).join('');
      m.querySelectorAll('[data-c]').forEach(c => c.onclick = async () => {
        try { await setDoc(doc(X.db, 'barcodes', bkey(code)), { c: c.dataset.c, code: String(code).trim(), by: S.user.name, at: Date.now() }); S.barcodes[bkey(code)] = c.dataset.c; (S.barcodeCodes ||= {})[bkey(code)] = String(code).trim(); }
        catch (e) { return toast('Could not save the barcode'); }
        m.remove(); toast(addItem(c.dataset.c) + ' • barcode saved');
      });
    };
  }
  function addItem(code) {
    const p = by(code); if (!p) return `Not found: ${code}`;
    if (!R && p.q <= 0) return `${p.c} is OUT OF STOCK`;
    const ex = cart.items.find(i => i.c === p.c);
    if (ex) { if (!R && ex.q >= ex.stock) return `Only ${ex.stock} in stock`; ex.q += 1; }
    else cart.items.push({ c: p.c, en: p.en, ar: p.ar, pk: p.pk || 1, stock: p.q, list: p.p, p: r3(p.p * (1 - (cart.cust.disc || 0) / 100)), q: 1, note: '' });
    saveCart(kind); draw(); return `Added ${p.c}`;
  }
  function searchItem() {
    const m = X.modal(`<h3>Add item</h3><input id="si" placeholder="Item number or name…" autofocus><div id="sr"></div><button class="btn l" id="cl">Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelector('#si').oninput = e => {
      const t = e.target.value.trim().toLowerCase(); if (t.length < 2) { m.querySelector('#sr').innerHTML = ''; return; }
      const L = P.filter(p => p.c.toLowerCase().includes(t) || (p.en || '').toLowerCase().includes(t) || (p.ar || '').includes(t)).slice(0, 25);
      m.querySelector('#sr').innerHTML = L.map(p => `<div class="card" data-c="${esc(p.c)}"><div class="row"><b>${esc(p.c)}</b><span class="pill ${p.q <= 0 ? 'r' : 'g'}">${p.q <= 0 ? 'OUT' : p.q}</span></div><div class="sub">${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div></div>`).join('');
      m.querySelectorAll('[data-c]').forEach(c => c.onclick = () => { const r = addItem(c.dataset.c); toast(r); if (r.startsWith('Added')) m.remove(); });
    };
  }
  async function create() {
    const btn = $('#cr'); btn.disabled = true; btn.textContent = 'Creating…';
    const items = cart.items.map(i => ({ c: i.c, en: i.en, pk: i.pk, q: i.q, p: i.p, list: i.list }));
    const now = new Date(); let inv;
    try {
      await runTransaction(X.db, async tx => {
        const cref = doc(X.db, 'counters', 'invoice'), cs = await tx.get(cref), no = (cs.exists() ? cs.data().n : 0) + 1;
        const refs = items.map(i => doc(X.db, 'products', i.c)), snaps = [];
        for (const r of refs) snaps.push(await tx.get(r));
        snaps.forEach((s, i) => { if (!s.exists() || s.data().q < items[i].q) throw new Error('STOCK:' + items[i].c); });
        snaps.forEach((s, i) => tx.update(refs[i], { q: s.data().q - items[i].q }));
        tx.set(cref, { n: no });
        inv = { no: 'INV-' + pad(no), cust: cart.cust.id, custEn: cart.cust.en, custAr: cart.cust.ar || '', disc: cart.cust.disc || 0, items,
          total: r3(items.reduce((a, i) => a + i.q * i.p, 0)), pieces: items.reduce((a, i) => a + i.q, 0), cartons: items.reduce((a, i) => a + i.q / i.pk, 0),
          by: S.user.uid, byName: S.user.name, created: now.getTime(), date: iso(now), delivery: cart.delivery || iso(now), note: cart.note || '', status: 'active', deliveryStatus: 'pending' };
        tx.set(doc(X.db, 'invoices', inv.no), inv);
      });
    } catch (e) {
      btn.disabled = false; btn.textContent = 'Create invoice';
      if (String(e.message).startsWith('STOCK:')) { await loadProducts(true); return alert('Not enough stock for ' + e.message.slice(6) + '. Stock was updated; please check quantities.'); }
      console.error(e); return toast('Could not create invoice. Try again.');
    }
    S.products = null; S.cart = null; saveCart('invoice'); logEvent('invoice', { ref: inv.no, cust: inv.cust, custEn: inv.custEn, total: inv.total }); done(inv);
  }
  async function createReturn() {
    const btn = $('#cr'); btn.disabled = true; btn.textContent = 'Saving…';
    const items = cart.items.map(i => ({ c: i.c, en: i.en, pk: i.pk, q: i.q, p: i.p, list: i.list, damaged: isDamaged(i.note) }));
    const now = new Date(), good = items.filter(i => !i.damaged); let ret;
    try {
      await runTransaction(X.db, async tx => {
        const cref = doc(X.db, 'counters', 'return'), cs = await tx.get(cref), no = (cs.exists() ? cs.data().n : 0) + 1;
        const refs = good.map(i => doc(X.db, 'products', i.c)), snaps = [];
        for (const r of refs) snaps.push(await tx.get(r));
        snaps.forEach((s, i) => { if (s.exists()) tx.update(refs[i], { q: s.data().q + good[i].q }); });
        tx.set(cref, { n: no });
        ret = { no: 'RET-' + pad(no), cust: cart.cust.id, custEn: cart.cust.en, custAr: cart.cust.ar || '', disc: cart.cust.disc || 0, items,
          total: r3(items.reduce((a, i) => a + i.q * i.p, 0)), pieces: items.reduce((a, i) => a + i.q, 0), cartons: items.reduce((a, i) => a + i.q / i.pk, 0),
          restocked: good.reduce((a, i) => a + i.q, 0), damagedPcs: items.filter(i => i.damaged).reduce((a, i) => a + i.q, 0),
          by: S.user.uid, byName: S.user.name, created: now.getTime(), date: iso(now), note: cart.note || '', status: 'active' };
        tx.set(doc(X.db, 'returns', ret.no), ret);
      });
    } catch (e) { console.error(e); btn.disabled = false; btn.textContent = 'Create return'; return toast('Could not save the return. Try again.'); }
    S.products = null; S.rcart = null; saveCart('return'); logEvent('return', { ref: ret.no, cust: ret.cust, custEn: ret.custEn, total: ret.total }); doneReturn(ret);
  }
  draw();
}
function done(inv) {
  const { top, esc, go, kwd, toast } = X;
  const fname = `${safe(inv.custEn)} ${dmy(new Date(inv.created))}.xlsx`;
  $('#app').innerHTML = top('Invoice created ✓') + `<div class="wrap"><div class="card" style="text-align:center"><div style="font-size:44px">✅</div><h2>${esc(inv.no)}</h2>
    <div>${esc(inv.custEn)}</div><div class="ar sub">${esc(inv.custAr)}</div><div style="font-size:24px;font-weight:800;margin-top:8px">${kwd(inv.total)} KWD</div>
    <div class="sub">${inv.pieces} pieces • ${inv.cartons.toFixed(2)} cartons</div></div>
    <button class="btn g" id="sh">📤 Share (WhatsApp / Email)</button><button class="btn" id="dl">⬇️ Download Excel</button><button class="btn l" id="nw">New invoice</button><button class="btn l" id="hm">Home</button></div>`;
  $('#bk').onclick = () => go('home');
  const run = async mode => { try { await sendFile(await buildExcel(inv.items, fname), fname, mode); } catch (e) { console.error(e); toast('Could not build the Excel file'); } };
  $('#sh').onclick = () => run('share'); $('#dl').onclick = () => run('download');
  $('#nw').onclick = () => go('invoice'); $('#hm').onclick = () => go('home');
}

/* ---------- Recent invoices ---------- */
async function recentList() {
  const { S, esc, kwd, toast } = X;
  let L;
  try { L = (await getDocs(query(collection(X.db, 'invoices'), orderBy('created', 'desc'), limit(80)))).docs.map(d => ({ id: d.id, ...d.data() })); }
  catch (e) { return ($('#rl').innerHTML = '<div class="card">Could not load invoices.</div>'); }
  if (S.user.role === 'sales') L = L.filter(i => i.by === S.user.uid);
  $('#rl').innerHTML = L.length ? L.map(i => `<div class="card" data-id="${i.id}"><div class="row"><b>${esc(i.no)}</b><span class="pill ${i.status === 'cancelled' ? 'r' : 'g'}">${i.status === 'cancelled' ? 'CANCELLED' : kwd(i.total)}</span></div>
    <div>${esc(i.custEn)}</div><div class="sub">${esc(i.date)} • by ${esc(i.byName)} • delivery ${esc(i.delivery)}</div></div>`).join('') : '<div class="card sub">No invoices yet.</div>';
  document.querySelectorAll('#rl .card[data-id]').forEach(c => c.onclick = () => openInvoice(L.find(i => i.id === c.dataset.id)));
}

function openInvoice(inv) {
  const { S, esc, kwd, toast } = X;
  const fname = `${safe(inv.custEn)} ${dmy(new Date(inv.created))}.xlsx`, admin = S.user.role === 'admin';
  const m = X.modal(`<h3>${esc(inv.no)} ${inv.status === 'cancelled' ? '<span class="pill r">CANCELLED</span>' : ''}</h3><div>${esc(inv.custEn)}</div><div class="ar sub">${esc(inv.custAr)}</div>
    <div class="sub">${esc(inv.date)} • by ${esc(inv.byName)} • delivery ${esc(inv.delivery)}</div>
    ${inv.items.map(i => `<div class="row"><span>${esc(i.c)} × ${i.q}</span><span>${kwd(i.q * i.p)}</span></div>`).join('')}
    <div class="row" style="margin-top:8px"><b>Total</b><b>${kwd(inv.total)} KWD</b></div>
    <button class="btn g" id="sh">📤 Share Excel</button><button class="btn" id="dl">⬇️ Download Excel</button>
    ${admin && inv.status === 'active' ? '<button class="btn r" id="cx">Cancel this invoice (stock returns)</button>' : ''}<button class="btn l" id="cl">Close</button>`);
  const run = async mode => { try { await sendFile(await buildExcel(inv.items, fname), fname, mode); } catch (e) { toast('Could not build the Excel file'); } };
  m.querySelector('#sh').onclick = () => run('share'); m.querySelector('#dl').onclick = () => run('download'); m.querySelector('#cl').onclick = () => m.remove();
  const cx = m.querySelector('#cx');
  if (cx) cx.onclick = async () => {
    if (!confirm(`Cancel ${inv.no}? The items go back into stock.`)) return;
    try {
      await runTransaction(X.db, async tx => {
        const ir = doc(X.db, 'invoices', inv.id), is = await tx.get(ir); if (is.data().status !== 'active') throw new Error('done');
        const refs = inv.items.map(i => doc(X.db, 'products', i.c)), snaps = [];
        for (const r of refs) snaps.push(await tx.get(r));
        snaps.forEach((s, i) => { if (s.exists()) tx.update(refs[i], { q: s.data().q + inv.items[i].q }); });
        tx.update(ir, { status: 'cancelled', cancelledBy: S.user.name, cancelledAt: Date.now() });
      });
      logEvent('invoice_cancel', { ref: inv.no, cust: inv.cust, custEn: inv.custEn }); X.S.products = null; m.remove(); toast('Invoice cancelled, stock returned'); X.go('invoice');
    } catch (e) { toast('Could not cancel'); }
  };
}

/* ---------- Returns: finished screen, recent list ---------- */
function doneReturn(ret) {
  const { top, esc, go, kwd, toast } = X;
  const fname = `${safe(ret.custEn)} RETURN ${dmy(new Date(ret.created))}.xlsx`;
  $('#app').innerHTML = top('Return saved ✓') + `<div class="wrap"><div class="card" style="text-align:center"><div style="font-size:44px">↩️</div><h2>${esc(ret.no)}</h2>
    <div>${esc(ret.custEn)}</div><div class="ar sub">${esc(ret.custAr)}</div><div style="font-size:24px;font-weight:800;margin-top:8px">${kwd(ret.total)} KWD credit</div>
    <div class="sub">${ret.pieces} pieces • back in stock: ${ret.restocked} • damaged: ${ret.damagedPcs}</div></div>
    <button class="btn g" id="sh">📤 Share (WhatsApp / Email)</button><button class="btn" id="dl">⬇️ Download Excel</button><button class="btn l" id="nw">New return</button><button class="btn l" id="hm">Home</button></div>`;
  $('#bk').onclick = () => go('home');
  const run = async mode => { try { await sendFile(await buildExcel(ret.items, fname), fname, mode); } catch (e) { console.error(e); toast('Could not build the Excel file'); } };
  $('#sh').onclick = () => run('share'); $('#dl').onclick = () => run('download');
  $('#nw').onclick = () => go('return'); $('#hm').onclick = () => go('home');
}
async function recentReturns() {
  const { S, esc, kwd, toast } = X;
  let L;
  try { L = (await getDocs(query(collection(X.db, 'returns'), orderBy('created', 'desc'), limit(80)))).docs.map(d => ({ id: d.id, ...d.data() })); }
  catch (e) { return ($('#rl').innerHTML = '<div class="card">Could not load returns.</div>'); }
  if (S.user.role === 'sales') L = L.filter(i => i.by === S.user.uid);
  $('#rl').innerHTML = L.length ? L.map(i => `<div class="card" data-id="${i.id}"><div class="row"><b>${esc(i.no)}</b><span class="pill g">${kwd(i.total)}</span></div>
    <div>${esc(i.custEn)}</div><div class="sub">${esc(i.date)} • by ${esc(i.byName)} • damaged ${i.damagedPcs}</div></div>`).join('') : '<div class="card sub">No returns yet.</div>';
  document.querySelectorAll('#rl .card[data-id]').forEach(c => c.onclick = () => {
    const r = L.find(i => i.id === c.dataset.id), fname = `${safe(r.custEn)} RETURN ${dmy(new Date(r.created))}.xlsx`;
    const m = X.modal(`<h3>${esc(r.no)}</h3><div>${esc(r.custEn)}</div><div class="sub">${esc(r.date)} • by ${esc(r.byName)}</div>
      ${r.items.map(i => `<div class="row"><span>${esc(i.c)} × ${i.q} ${i.damaged ? '<span class="pill r">damaged</span>' : ''}</span><span>${kwd(i.q * i.p)}</span></div>`).join('')}
      <div class="row" style="margin-top:8px"><b>Credit</b><b>${kwd(r.total)} KWD</b></div>
      <button class="btn g" id="sh">📤 Share Excel</button><button class="btn" id="dl">⬇️ Download Excel</button><button class="btn l" id="cl">Close</button>`);
    const run = async mode => { try { await sendFile(await buildExcel(r.items, fname), fname, mode); } catch (e) { toast('Could not build the Excel file'); } };
    m.querySelector('#sh').onclick = () => run('share'); m.querySelector('#dl').onclick = () => run('download'); m.querySelector('#cl').onclick = () => m.remove();
  });
}
