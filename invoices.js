import { emailInvoice, emailReturn } from './mailer.js';
import { logEvent } from './activity.js';
import { doc, getDocs, setDoc, collection, runTransaction, query, orderBy, limit } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

let X; // shared context from app.js
export function initInvoices(ctx) { X = ctx; }

const $ = s => document.querySelector(s);
const nm = x => (x.custEn || '') + (x.branch ? ' – ' + x.branch : '');
const brKey = id => 'kw_br_' + id;
const brList = id => { try { return JSON.parse(localStorage.getItem(brKey(id)) || '[]'); } catch (e) { return []; } };
const brSave = (id, b) => { b = (b || '').trim(); if (!b) return; try { const L = brList(id).filter(x => x !== b); L.unshift(b); localStorage.setItem(brKey(id), JSON.stringify(L.slice(0, 15))); } catch (e) {} };
const r3 = n => Math.round(n * 1000) / 1000;
const pad = n => String(n).padStart(5, '0');
const dmy = d => `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const safe = s => String(s).replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
const pid = c => (((X.S.products || []).find(p => p.c === c)) || {}).id || c;   // product document id for an item code (they differ for renamed items)
const lavOf = c => { const p = (X.S.products || []).find(p => p.c === c); return !!p && X.catOf(p) === 'lav'; };
const cf = i => (i.lav ?? lavOf(i.c)) ? 1 : (i.pk || 1);   // LAV: every unit added is already one carton (not divided by packing)
const loaded = {};
function loadScript(url) {
  return loaded[url] ||= new Promise((res, rej) => { const s = document.createElement('script'); s.src = url; s.onload = res; s.onerror = () => rej(new Error('load ' + url)); document.head.appendChild(s); });
}
const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const ZXING = 'https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js';

/* ---------- Excel (shared by invoices and returns) ---------- */
// Discount is taken off the invoice TOTAL (not per item): subtotal at list prices, minus discount %, equals the final price.
export const calc = (items, d) => { const sub = r3(items.reduce((a, i) => a + i.q * i.p, 0)), da = r3(sub * (d || 0) / 100); return { sub, da, net: r3(sub - da) }; };
const dOf = x => (x && x.dm === 'end' ? x.disc || 0 : 0);
export async function buildExcel(items, name, disc) {
  await loadScript(EXCELJS);
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Invoice');
  ws.columns = [{ width: 24 }, { width: 14 }, { width: 14 }];
  const h = ws.addRow(['Item No.', 'Quantity', 'Price']); h.font = { bold: true };
  items.forEach(it => { const r = ws.addRow([String(it.c), it.q, it.p]); r.getCell(3).numFmt = '0.000'; r.getCell(1).numFmt = '@'; });
  const n = items.length, last = n + 1, t = n + 3;
  const qty = items.reduce((a, i) => a + i.q, 0), price = r3(items.reduce((a, i) => a + i.q * i.p, 0)), cart = items.reduce((a, i) => a + i.q / cf(i), 0);
  const put = (row, label, formula, result, fmt) => { ws.getCell(`A${row}`).value = label; ws.getCell(`A${row}`).font = { bold: true }; const c = ws.getCell(`B${row}`); c.value = { formula, result }; c.font = { bold: true }; if (fmt) c.numFmt = fmt; };
  put(t, 'Final Quantity', `SUM(B2:B${last})`, qty);
  const d = disc || 0;
  if (d) {
    put(t + 1, 'Subtotal (KWD)', `SUMPRODUCT(B2:B${last},C2:C${last})`, price, '0.000');
    put(t + 2, `Discount ${d}% (KWD)`, `-ROUND(B${t + 1}*${d}/100,3)`, -r3(price * d / 100), '0.000');
    put(t + 3, 'Final Price (KWD)', `B${t + 1}+B${t + 2}`, r3(price - r3(price * d / 100)), '0.000');
    put(t + 4, 'Cartons', items.map((it, i) => `B${i + 2}/${cf(it)}`).join('+'), cart, '0.##');
  } else {
    put(t + 1, 'Final Price (KWD)', `SUMPRODUCT(B2:B${last},C2:C${last})`, price, '0.000');
    put(t + 2, 'Cartons', items.map((it, i) => `B${i + 2}/${cf(it)}`).join('+'), cart, '0.##');
  }
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
export async function buildStockExcel(products, inv, flag) {
  await loadScript(EXCELJS);
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Stock');
  ws.columns = [{ width: 18 }, { width: 30 }, { width: 36 }, { width: 10 }, { width: 12 }, { width: 10 }, { width: 22 }, { width: 40 }, { width: 44 }];
  const h = ws.addRow(['Item Number', 'Arabic Name', 'English Name', 'Packing', 'Stock Qty', 'Price (KWD)', 'Barcode (item number)', 'Added Barcodes', 'Status']); h.font = { bold: true };
  [...products].sort((a, b) => String(a.c).localeCompare(String(b.c))).forEach(p => {
    const r = ws.addRow([p.c, p.ar || '', p.en || '', p.pk, p.q, p.p, p.c, (inv[p.c] || []).join(', '), flag ? flag(p) : '']);
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

/* ---------- Live scanner: stays open, reads barcodes AND printed item numbers automatically ---------- */
function beep(ok) {
  try { navigator.vibrate && navigator.vibrate(ok ? 60 : [80, 60, 80]); } catch (e) {}
  try {
    const a = new (window.AudioContext || window.webkitAudioContext)(), o = a.createOscillator(), g = a.createGain();
    o.frequency.value = ok ? 1100 : 300; g.gain.value = 0.08; o.connect(g); g.connect(a.destination); o.start(); setTimeout(() => { o.stop(); a.close(); }, ok ? 110 : 260);
  } catch (e) {}
}
// opts: onCode(code)->{ok,msg,c?,unknown?} | onText(text)->{ok,msg,c?}|null | setQty(c,n)->n | getQty(c) | onDone()
export function scanner(opts) {
  const ov = document.createElement('div'); ov.className = 'scan-ov';
  ov.innerHTML = `<div class="scan-box"><video id="vd" playsinline muted></video><div class="scan-frame"></div></div>
    <div id="sm" class="scan-msg">Starting camera…</div>
    <div id="lc" class="scan-last"></div>
    <div class="row" style="gap:8px;margin:6px 0"><button class="btn l" id="tq" style="display:none">🔦 Light</button><input id="zm" type="range" style="display:none;flex:1"></div>
    <div class="row" style="gap:8px"><input id="mi" placeholder="Or type item number" style="flex:1"><button class="btn l" id="ma">Add</button></div>
    <button class="btn g" id="cl">Done</button>`;
  document.body.appendChild(ov);
  const $$ = s => ov.querySelector(s);
  let reader, stopped = false, paused = false, track, last = '', lastT = 0, pend = '', pendN = 0, ocrBusy = false, ocrKey = '', ocrN = 0, ocrDone = '', lastRead = 0, cur = null;
  const stop = () => { if (stopped) return; stopped = true; try { track && track.applyConstraints({ advanced: [{ torch: false }] }); } catch (e) {} try { reader && reader.reset(); } catch (e) {} ov.remove(); opts.onDone && opts.onDone(); };
  $$('#cl').onclick = stop;
  const say = (t, ok) => { const s = $$('#sm'); s.textContent = t; s.className = 'scan-msg' + (ok === true ? ' ok' : ok === false ? ' bad' : ''); };
  const showLast = r => {
    cur = r.c || null; const box = $$('#lc');
    if (!cur) { box.innerHTML = ''; return; }
    const draw = () => { const sh = opts.shelfMode && opts.getShelf(cur);
      box.innerHTML = `<div class="row">${X.img(cur)}<b style="flex:1;margin-left:8px">${X.esc(cur)}</b>${opts.shelfMode ? `<button class="btn ${sh ? 'g' : 'l'}" id="shb" style="margin:0;width:auto">✔ متوفر${sh ? ' ✓' : ''}</button>` : ''}</div>
        <div class="row" style="gap:6px;margin-top:6px">${opts.shelfMode ? '<span>🛒 Order</span>' : '<span></span>'}<span class="row" style="gap:6px"><button class="btn l" data-d="-1">−</button><input id="qn" type="number" inputmode="numeric" value="${opts.getQty(cur)}" style="width:64px;text-align:center"><button class="btn l" data-d="1">+</button></span></div>`;
      const sb = box.querySelector('#shb'); if (sb) sb.onclick = () => { opts.toggleShelf(cur); draw(); };
      box.querySelectorAll('[data-d]').forEach(b => b.onclick = () => { const n = opts.setQty(cur, opts.getQty(cur) + +b.dataset.d); say(`${cur} → ${n}`, true); draw(); });
      box.querySelector('#qn').onchange = e => { const n = opts.setQty(cur, +e.target.value || (opts.shelfMode ? 0 : 1)); e.target.value = n; }; };
    draw();
  };
  const handle = r => {
    if (!r) return false; lastRead = Date.now();
    beep(r.ok); say(r.msg, r.ok); showLast(r);
    if (r.unknown) { paused = true; stop(); setTimeout(() => opts.onUnknown(r.unknown), 50); }
    return true;
  };
  $$('#ma').onclick = () => { const v = $$('#mi').value.trim(); if (!v) return; $$('#mi').value = ''; const r = opts.onCode(v); r.unknown ? (say('Not found: ' + v, false), beep(false)) : handle(r); };
  (async () => {
    try {
      await loadScript(ZXING);
      const hints = new Map(), F = ZXing.BarcodeFormat;
      hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E, F.CODE_128, F.CODE_39]);
      hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
      const mfr = new ZXing.MultiFormatReader(); mfr.setHints(hints);
      const v = $$('#vd');
      say('Point at the barcode or the printed item number');
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } } });
      v.srcObject = stream; await v.play();
      reader = { reset: () => stream.getTracks().forEach(t => t.stop()) };
      const BD = window.BarcodeDetector ? new BarcodeDetector({ formats: ['ean_13', 'ean_8', 'upc_a', 'upc_e', 'code_128', 'code_39'] }) : null;
      const cv = document.createElement('canvas'), cx = cv.getContext('2d', { willReadFrequently: true });
      // Shiny packaging reflects light when the phone is straight on, which blinds the reader. So every frame is tried several ways:
      // as is, with the bright glare flattened, and with the contrast stretched.
      const enhance = mode => {
        const d = cx.getImageData(0, 0, cv.width, cv.height), p = d.data, n = p.length / 4, hist = new Uint32Array(256);
        for (let i = 0; i < p.length; i += 4) { const g = (p[i] * 0.3 + p[i + 1] * 0.59 + p[i + 2] * 0.11) | 0; p[i] = g; hist[g]++; }
        let lo = 0, hi = 255, acc = 0; while (lo < 255 && (acc += hist[lo]) < n * 0.02) lo++; acc = 0; while (hi > 0 && (acc += hist[hi]) < n * 0.02) hi--;
        const cap = mode === 'glare' ? Math.max(lo + 20, hi - (hi - lo) * 0.25) : hi, r = Math.max(20, cap - lo);
        for (let i = 0; i < p.length; i += 4) { let g = Math.min(255, Math.max(0, ((Math.min(p[i], cap) - lo) * 255 / r) | 0)); p[i] = p[i + 1] = p[i + 2] = g; p[i + 3] = 255; }
        cx.putImageData(d, 0, 0);
      };
      const tryDecode = () => { for (const B of [ZXing.HybridBinarizer, ZXing.GlobalHistogramBinarizer]) { try { return mfr.decodeWithState(new ZXing.BinaryBitmap(new B(new ZXing.HTMLCanvasElementLuminanceSource(cv)))); } catch (e) {} } return null; };
      let frame = 0;
      const scan = async () => {
        while (!stopped) {
          await new Promise(r => setTimeout(r, 70));
          if (paused || !v.videoWidth) continue;
          frame++;
          const w = v.videoWidth, h = v.videoHeight, cw = w, ch = h * 0.7, k = Math.min(1, 1280 / cw);
          cv.width = Math.round(cw * k); cv.height = Math.round(ch * k);
          cx.drawImage(v, 0, (h - ch) / 2, cw, ch, 0, 0, cv.width, cv.height);
          let t = null, strong = false;
          if (BD) { try { const r = await BD.detect(cv); if (r.length) { t = r[0].rawValue; strong = /^(ean_13|ean_8|upc_a|upc_e)$/.test(r[0].format); } } catch (e) {} }
          const modes = [null, 'glare', 'stretch'];
          for (const m of (t ? [] : modes)) {
            if (m) enhance(m);
            const r = tryDecode(); if (r) { t = r.getText(); const f = r.getBarcodeFormat(), F2 = ZXing.BarcodeFormat; strong = f === F2.EAN_13 || f === F2.EAN_8 || f === F2.UPC_A || f === F2.UPC_E; break; }
            if (m !== 'stretch') cx.drawImage(v, 0, (h - ch) / 2, cw, ch, 0, 0, cv.width, cv.height);
          }
          if (!t) continue;
          const now = Date.now();
          if (!strong) { if (t !== pend) { pend = t; pendN = 1; continue; } if (++pendN < 2) continue; }
          if (t === last && now - lastT < 1800) { lastT = now; continue; } last = t; lastT = now;
          handle(opts.onCode(t));
        }
      };
      scan();
      // light + zoom when the phone supports them
      setTimeout(() => {
        try {
          track = v.srcObject.getVideoTracks()[0]; const cap = track.getCapabilities ? track.getCapabilities() : {};
          if (cap.torch) { const b = $$('#tq'); b.style.display = ''; let on = false; b.onclick = () => { on = !on; track.applyConstraints({ advanced: [{ torch: on }] }).catch(() => {}); }; }
          if (cap.zoom) { const z = $$('#zm'); z.style.display = ''; z.min = cap.zoom.min; z.max = Math.min(cap.zoom.max, 5); z.step = cap.zoom.step || 0.1; z.value = cap.zoom.min; z.oninput = () => track.applyConstraints({ advanced: [{ zoom: +z.value }] }).catch(() => {}); }
          if (cap.focusMode && cap.focusMode.includes('continuous')) track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(() => {});
        } catch (e) {}
      }, 1200);
      // automatic reading of printed item numbers (no extra button): tries the middle of the picture every ~1 second
      if (opts.onText) {
        loadScript(TESS).catch(() => {});
        const loop = async () => {
          while (!stopped) {
            await new Promise(r => setTimeout(r, 900));
            if (stopped || paused || ocrBusy || !v.videoWidth || Date.now() - lastRead < 1500) continue;
            ocrBusy = true;
            try {
              const w = v.videoWidth, h = v.videoHeight, cw = w * 0.9, ch = h * 0.5, c = document.createElement('canvas'), k = Math.min(1, 1100 / cw);
              c.width = Math.round(cw * k); c.height = Math.round(ch * k);
              c.getContext('2d').drawImage(v, (w - cw) / 2, (h - ch) / 2, cw, ch, 0, 0, c.width, c.height);
              const r = opts.onText(await ocrText(c, 1100), true);
              if (r && r.key && r.key === ocrDone) { /* same label still in view */ }
              else if (r && r.key) { ocrDone = ''; if (r.key === ocrKey) ocrN++; else { ocrKey = r.key; ocrN = 1; } if (ocrN >= 2) { ocrDone = r.key; ocrKey = ''; ocrN = 0; if (!stopped && !paused) handle(opts.onText(null, false)); } }
              else { ocrKey = ''; ocrN = 0; ocrDone = ''; }
            } catch (e) { console.error(e); }
            ocrBusy = false;
          }
        };
        loop();
      }
    } catch (e) { say('Camera not available. Allow camera access, or type the item number below.', false); }
  })();
  return stop;
}


/* ---------- Read a printed item number from a photo (OCR) ---------- */
const TESS = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
let ocrWorker;
async function prepImage(file, max = 2200) {
  const img = await createImageBitmap(file), k = Math.min(max / Math.max(img.width, img.height), 2.5);
  const c = document.createElement('canvas'); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
  const x = c.getContext('2d'); x.drawImage(img, 0, 0, c.width, c.height);
  const d = x.getImageData(0, 0, c.width, c.height), g = new Uint8Array(c.width * c.height); let lo = 255, hi = 0;
  for (let i = 0, j = 0; i < d.data.length; i += 4, j++) { const v = (d.data[i] * 0.3 + d.data[i + 1] * 0.59 + d.data[i + 2] * 0.11) | 0; g[j] = v; if (v < lo) lo = v; if (v > hi) hi = v; }
  const r = Math.max(1, hi - lo);
  for (let i = 0, j = 0; i < d.data.length; i += 4, j++) { const v = Math.min(255, Math.max(0, ((g[j] - lo) * 255 / r) | 0)); d.data[i] = d.data[i + 1] = d.data[i + 2] = v; }
  x.putImageData(d, 0, 0); return c;
}
async function ocrText(file, max) {
  await loadScript(TESS);
  ocrWorker ||= await Tesseract.createWorker('eng');
  await ocrWorker.setParameters({ tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-', tessedit_pageseg_mode: '11' });
  return (await ocrWorker.recognize(await prepImage(file, max))).data.text;
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
function restoreCart(kind) { const [f, k] = KEYS[kind]; if (X.S[f]) return; try { X.S[f] = JSON.parse(localStorage.getItem(k) || 'null'); (X.S[f]?.items || []).forEach(i => { if (i.list != null) i.p = i.list; }); } catch (e) {} }
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
  const R = kind === 'return', F = KEYS[kind][0], TT = R ? 'Return • مرتجع' : 'Shelf Check • جرد الرف';
  $('#app').innerHTML = top(TT) + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  const C = await loadCustomers(); let q = '', tab = 'new';
  const draw = () => {
    const t = q.trim().toLowerCase();
    const L = C.filter(c => (c.active !== false || t) && (!t || c.id.includes(t) || (c.en || '').toLowerCase().includes(t) || (c.ar || '').includes(t))).sort((a, b) => (a.en || '').localeCompare(b.en || ''));
    $('#app').innerHTML = top(TT) + `<div class="wrap">
      <div class="tabs"><button data-t="new" class="${tab === 'new' ? 'on' : ''}">New</button><button data-t="recent" class="${tab === 'recent' ? 'on' : ''}">${R ? 'Recent returns' : 'Invoices'}</button>${R ? '' : `<button data-t="visits" class="${tab === 'visits' ? 'on' : ''}">Visits</button>`}</div>
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
    } else (R ? recentReturns() : tab === 'visits' ? recentVisits() : recentList());
  };
  draw();
}

async function cartView(kind) {
  const { S, top, esc, go, kwd, toast, loadProducts } = X;
  const R = kind === 'return', F = KEYS[kind][0], cart = S[F]; await loadProducts();
  await X.loadBarcodes();
  const bkey = c => String(c).trim().replace(/\//g, '_');
  const P = S.products, by = id => { const k = String(id).trim(); let p = P.find(p => p.id === k || p.c.toLowerCase() === k.toLowerCase()); if (!p && S.barcodes[bkey(k)]) p = P.find(p => p.c === S.barcodes[bkey(k)]); return p; };
  const total = () => calc(cart.items, cart.cust.disc).net, sub = () => calc(cart.items, cart.cust.disc);
  const draw = () => {
    const pcs = cart.items.reduce((a, i) => a + i.q, 0), ctn = cart.items.reduce((a, i) => a + i.q / cf(i), 0);
    $('#app').innerHTML = top((R ? 'Return • ' : 'Visit • ') + esc(cart.cust.en)) + `<div class="wrap">
      ${cart.cust.disc ? `<div class="pill g" style="margin-bottom:8px">Customer discount ${cart.cust.disc}%</div>` : ''}
      <div class="grid" style="margin-bottom:10px"><button class="btn" id="sc">📷 Scan</button><button class="btn l" id="se">🔎 Search item</button></div>
      <button class="btn l" id="ph" style="margin-top:0">${R ? '🖼 Photo: barcode or printed item number' : '🖼 Read barcode from photo'}</button><input type="file" id="pf" accept="image/*" multiple style="display:none">
      ${cart.items.length === 0 ? '<div class="card sub">No items yet. Scan a barcode or search for an item.</div>' : ''}
      ${R ? cart.items.map((it, i) => `<div class="card pc">${X.img(it.c)}<div class="pcb"><div class="row"><b>${esc(it.c)}</b><button class="pill r" data-rm="${i}" style="border:0">Remove</button></div>
        <div class="sub">${esc(it.en)}</div><div class="ar sub">${esc(it.ar)}</div>
        ${R ? '' : `<label style="display:flex;gap:10px;align-items:center;margin:6px 0"><input type="checkbox" data-sh="${i}" ${it.shelf ? 'checked' : ''} style="width:24px;height:24px;margin:0"> ✔ Available on his shelf • موجود عنده</label>`}
        <div class="row" style="margin-top:6px"><div style="flex:1"><label>${R ? 'Pieces returned' : 'Pieces to order (stock ' + it.stock + ')'}</label><input type="number" inputmode="numeric" data-q="${i}" value="${it.q}"></div>
        <button class="btn l" style="width:auto;margin:20px 0 0" data-ctn="${i}">+1 carton<br><span class="sub">(${it.pk} pcs)</span></button></div>
        ${R ? `<div class="row"><input data-nt="${i}" placeholder="Note — write damaged / تالف / K if broken" value="${esc(it.note || '')}" style="flex:1"><button class="btn l" style="width:auto;margin:6px 0" data-dm="${i}">🔨</button></div>${isDamaged(it.note) ? '<span class="pill r">Damaged — will NOT go back to stock</span>' : '<span class="pill g">Goes back to stock</span>'}` : ''}
        <div class="sub">${kwd(it.p)} KWD each • ${(it.q / cf(it)).toFixed(2)} cartons • <b>${kwd(it.q * it.p)}</b></div></div>`).join('') : `${(() => { const O = cart.items.map((it, i) => ({ it, i })).filter(x => x.it.q > 0), Sh = cart.items.map((it, i) => ({ it, i })).filter(x => x.it.shelf);
        return `${O.length ? `<table class="ot"><thead><tr><th>Item no.</th><th>Qty</th><th>Packing</th><th>Price</th></tr></thead><tbody>${O.map(({ it, i }) => `<tr><td><b>${esc(it.c)}</b>${it.shelf ? ' <span class="pill g">✔</span>' : ''}</td><td><input type="number" inputmode="numeric" data-q="${i}" value="${it.q}"><button class="pill m" data-ctn="${i}" style="border:0">+ctn</button></td><td>${it.pk}</td><td>${kwd(it.p)}</td></tr>`).join('')}</tbody>
          <tfoot><tr><td><b>TOTAL</b></td><td>${pcs}</td><td><b>${ctn.toFixed(2)}</b><br><span class="sub">cartons</span></td><td><b>${kwd(sub().sub)}</b><br><span class="sub">KWD</span></td></tr>${cart.cust.disc ? `<tr><td colspan="3">Discount ${cart.cust.disc}% • خصم</td><td>-${kwd(sub().da)}</td></tr><tr><td colspan="3"><b>FINAL • الصافي</b></td><td><b>${kwd(sub().net)}</b></td></tr>` : ''}</tfoot></table>` : ''}
        ${Sh.length ? `<h3 style="margin:14px 0 4px">✔ متوفر عنده (${Sh.length})</h3>${Sh.map(({ it, i }) => `<div class="row" style="padding:4px 0;border-bottom:1px solid #eee"><span><b>${esc(it.c)}</b> <span class="sub">${esc(it.en)}</span></span><label style="margin:0"><input type="checkbox" data-sh="${i}" checked style="width:22px;height:22px;margin:0"></label></div>`).join('')}` : ''}`; })()}`}
      ${R ? `<div class="card"><div class="row"><span>Pieces</span><b>${pcs}</b></div><div class="row"><span>Cartons</span><b>${ctn.toFixed(2)}</b></div>
        ${cart.cust.disc ? `<div class="row"><span>Subtotal</span><b>${kwd(sub().sub)}</b></div><div class="row"><span>Discount ${cart.cust.disc}%</span><b>-${kwd(sub().da)}</b></div>` : ''}<div class="row"><span>Total (KWD)</span><b style="font-size:22px">${kwd(total())}</b></div></div>` : ''}
      ${R ? '' : `<label>Delivery date • تاريخ التوصيل</label><input type="date" id="dd" value="${cart.delivery}">`}
      <label>Branch / city • الفرع / المنطقة (optional)</label><input id="br" list="brl" placeholder="e.g. Jahra, Salmiya" value="${esc(cart.branch || '')}"><datalist id="brl">${brList(cart.cust.id).map(b => `<option value="${esc(b)}">`).join('')}</datalist>
      <label>Note (optional)</label><input id="nt" value="${esc(cart.note)}">
      ${R ? '' : `<button class="btn" id="sv" ${cart.items.length ? '' : 'disabled style="opacity:.5"'}>💾 Save visit (shelf + draft order)</button>`}
      <button class="btn g" id="cr" ${(R ? cart.items.length : cart.items.some(i => i.q > 0)) ? '' : 'disabled style="opacity:.5"'}>${R ? 'Create return' : 'Create invoice from the order'}</button>
      <button class="btn l" id="cn">Cancel / start over</button></div>`;
    $('#bk').onclick = () => { go('home'); };
    document.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { cart.items.splice(+b.dataset.rm, 1); saveCart(kind); draw(); });
    document.querySelectorAll('[data-q]').forEach(inp => inp.onchange = () => {
      const it = cart.items[+inp.dataset.q]; let v = Math.max(R ? 1 : 0, Math.floor(+inp.value || (R ? 1 : 0)));
      if (!R && v > it.stock) { v = it.stock; toast(`Only ${it.stock} in stock`); } it.q = v; if (!R && !v && !it.shelf) cart.items.splice(+inp.dataset.q, 1); saveCart(kind); draw();
    });
    document.querySelectorAll('[data-ctn]').forEach(b => b.onclick = () => {
      const it = cart.items[+b.dataset.ctn]; const st = lavOf(it.c) ? 1 : it.pk; if (!R && it.q + st > it.stock) { toast(`Only ${it.stock} in stock`); it.q = it.stock; } else it.q += st; saveCart(kind); draw();
    });
    document.querySelectorAll('[data-sh]').forEach(c => c.onchange = () => { const it = cart.items[+c.dataset.sh]; it.shelf = c.checked; if (!it.shelf && !it.q) cart.items.splice(+c.dataset.sh, 1); saveCart(kind); draw(); });
    const sv = $('#sv'); if (sv) sv.onclick = saveOnly;
    const dd = $('#dd'); if (dd) dd.onchange = e => { cart.delivery = e.target.value; saveCart(kind); };
    document.querySelectorAll('[data-nt]').forEach(inp => inp.onchange = () => { cart.items[+inp.dataset.nt].note = inp.value; saveCart(kind); draw(); });
    document.querySelectorAll('[data-dm]').forEach(b => b.onclick = () => { cart.items[+b.dataset.dm].note = 'damaged'; saveCart(kind); draw(); });
    $('#nt').onchange = e => { cart.note = e.target.value; saveCart(kind); };
    $('#br').oninput = e => { cart.branch = e.target.value; saveCart(kind); };
    $('#cn').onclick = () => { if (!cart.items.length || confirm('Throw away this draft?')) { S[F] = null; saveCart(kind); go(R ? 'return' : 'invoice'); } };
    $('#sc').onclick = () => {
      let pendingHit = null;
      const res = c => { if (!R) { const p = by(c); return { ok: true, msg: `${p.c} — ${p.en || ''}`, c: p.c }; } const r = addItem(c), p = by(c); return { ok: r.startsWith('Added'), msg: r.startsWith('Added') ? `✔ ${p.c} — ${p.en || ''}` : r, c: p ? p.c : null }; };
      scanner({
        onCode: code => by(code) ? res(code) : { ok: false, unknown: code, msg: 'Unknown barcode ' + code },
        onText: (text, probe) => {
          if (probe) { const h = matchText(text, S.products); pendingHit = h.length === 1 ? h[0] : null; return pendingHit ? { key: pendingHit.c } : null; }
          return pendingHit ? res(pendingHit.c) : null;
        },
        onUnknown: code => linkBarcode(code),
        shelfMode: !R,
        getShelf: c => !!(cart.items.find(i => i.c === c) || {}).shelf,
        toggleShelf: c => toggleShelf(c),
        getQty: c => (cart.items.find(i => i.c === c) || {}).q || 0,
        setQty: (c, n) => { if (!R) return setOrder(c, n); const it = cart.items.find(i => i.c === c); if (!it) return 0; n = Math.max(1, Math.floor(n) || 1); it.q = n; saveCart(kind); draw(); return n; }
      });
    };
    $('#se').onclick = searchItem;
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
      m.querySelector('#sr').innerHTML = L.map(p => `<div class="card pc" data-c="${esc(p.c)}">${X.img(p.c)}<div class="pcb"><div class="row"><b>${esc(p.c)}</b><span class="pill ${p.q <= 0 ? 'r' : 'g'}">${p.q <= 0 ? 'OUT' : p.q}</span></div><div class="sub">${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div></div></div>`).join('');
      m.querySelectorAll('[data-c]').forEach(c => c.onclick = async () => {
        try { await setDoc(doc(X.db, 'barcodes', bkey(code)), { c: c.dataset.c, code: String(code).trim(), by: S.user.name, at: Date.now() }); S.barcodes[bkey(code)] = c.dataset.c; (S.barcodeCodes ||= {})[bkey(code)] = String(code).trim(); }
        catch (e) { return toast('Could not save the barcode'); }
        m.remove(); toast(addItem(c.dataset.c) + ' • barcode saved');
      });
    };
  }
  function ensure(c) {
    const p = by(c); if (!p) return null; let it = cart.items.find(i => i.c === p.c);
    if (!it) { it = { c: p.c, en: p.en, ar: p.ar, pk: p.pk || 1, stock: p.q, list: p.p, p: p.p, q: 0, shelf: false, note: '' }; cart.items.push(it); }
    return it;
  }
  const tidy = it => { if (it && !it.q && !it.shelf) { const k = cart.items.indexOf(it); if (k >= 0) cart.items.splice(k, 1); } };
  function toggleShelf(c) { const it = ensure(c); if (!it) return false; it.shelf = !it.shelf; const r = it.shelf; tidy(it); saveCart(kind); draw(); return r; }
  function setOrder(c, n) {
    const it = ensure(c); if (!it) return 0; n = Math.max(0, Math.floor(n) || 0);
    if (n > it.stock) { toast(it.stock > 0 ? `Only ${it.stock} in stock` : `${c} is OUT OF STOCK`); n = Math.max(0, it.stock); }
    it.q = n; tidy(it); saveCart(kind); draw(); return n;
  }
  function choose(c) {   // small window for one item: tick "on his shelf" and/or set pieces to order
    const p = by(c); if (!p) return;
    const m = X.modal(`<div id="ch"></div><button class="btn g" id="cl">Done</button>`);
    const d = () => { const it = cart.items.find(i => i.c === p.c) || {};
      m.querySelector('#ch').innerHTML = `<div class="card pc">${X.img(p.c)}<div class="pcb"><b>${esc(p.c)}</b><div class="sub">${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div><div class="sub">stock ${p.q}</div></div></div>
        <button class="btn ${it.shelf ? 'g' : 'l'}" data-s>✔ متوفر${it.shelf ? ' ✓' : ''}</button>
        <div class="row" style="gap:6px"><span>🛒 Order (pieces)</span><span class="row" style="gap:6px"><button class="btn l" data-o="-1">−</button><b style="min-width:30px;text-align:center">${it.q || 0}</b><button class="btn l" data-o="1">+</button>${lavOf(p.c) ? '' : `<button class="btn l" data-o="${p.pk || 1}">+${p.pk || 1}</button>`}</span></div>`;
      m.querySelector('[data-s]').onclick = () => { toggleShelf(p.c); d(); };
      m.querySelectorAll('[data-o]').forEach(b => b.onclick = () => { setOrder(p.c, ((cart.items.find(i => i.c === p.c) || {}).q || 0) + +b.dataset.o); d(); }); };
    m.querySelector('#cl').onclick = () => m.remove(); d();
  }
  function addItem(code, toggle) {
    const p = by(code); if (!p) return `Not found: ${code}`;
    const ex = cart.items.find(i => i.c === p.c);
    if (!R) { choose(p.c); return `Added ${p.c} — choose shelf / order`; }
    if (!R && p.q <= 0) return `${p.c} is OUT OF STOCK`;
    if (ex) { if (!R && ex.q >= ex.stock) return `Only ${ex.stock} in stock`; ex.q += 1; }
    else cart.items.push({ c: p.c, en: p.en, ar: p.ar, pk: p.pk || 1, stock: p.q, list: p.p, p: p.p, q: 1, shelf: false, note: '' });
    saveCart(kind); draw(); return `Added ${p.c}`;
  }
  function searchItem() {
    const PV = P.filter(p => !X.isHidden(p) && (R || p.q > 0));
    let cat = 'all', sub = 'all', t = '', lim = 40;
    const PV0 = P.filter(p => !X.isHidden(p));
    const m = X.modal(`<h3 style="margin:0 0 4px;font-size:15px">Add items</h3><input id="si" placeholder="Item number or name…" autofocus style="padding:8px 12px;font-size:16px;margin:2px 0">
      <div class="tabs cats" id="ct"></div><div class="tabs cats" id="cs"></div><div id="sr"></div><button class="btn g" id="cl">Done</button>`);
    const kinds = [['all', 'All'], ...X.CATLIST];
    const drawCats = () => { m.querySelector('#ct').innerHTML = kinds.map(([k, n]) => `<button data-k="${k}" class="${cat === k ? 'on' : ''}">${n}</button>`).join(''); m.querySelectorAll('#ct button').forEach(b => b.onclick = () => { cat = b.dataset.k; sub = 'all'; lim = 40; drawCats(); drawList(); }); drawSubs(); };
    const drawSubs = () => { const L = X.SUBS[cat]; m.querySelector('#cs').innerHTML = L ? [['all', 'All'], ...L].map(([k, n]) => `<button data-s="${k}" class="${sub === k ? 'on' : ''}">${n}</button>`).join('') : ''; m.querySelectorAll('#cs button').forEach(b => b.onclick = () => { sub = b.dataset.s; lim = 40; drawSubs(); drawList(); }); };
    const drawList = () => {
      const L = (!R && t.length >= 2 ? PV0 : PV).filter(p => (cat === 'all' || X.catOf(p) === cat) && (sub === 'all' || X.subOf(p) === sub) && (!t || p.c.toLowerCase().includes(t) || (p.en || '').toLowerCase().includes(t) || (p.ar || '').includes(t) || Object.keys(S.barcodes || {}).some(b => S.barcodes[b] === p.c && (((S.barcodeCodes || {})[b] || b).toLowerCase().includes(t)))));
      if (cat !== 'all' && X.SUBS[cat]) L.sort((x, y) => X.subRank(x) - X.subRank(y) || x.c.localeCompare(y.c));
      if (cat === 'all' && t.length < 2) { m.querySelector('#sr').innerHTML = '<div class="sub">Type to search, or pick a category to browse it.</div>'; return; }
      m.querySelector('#sr').innerHTML = L.slice(0, lim).map(p => { const ci = cart.items.find(i => i.c === p.c) || {}, n = ci.q || 0, sh = ci.shelf;
        if (!R) return `<div class="card pc cp" data-c="${esc(p.c)}">${X.img(p.c)}<div class="pcb"><div class="t1">${esc(p.c)} <span class="pill ${p.q <= 0 ? 'r' : 'g'}">${p.q <= 0 ? 'OUT' : p.q}</span></div><div class="sub">${esc(p.en)}</div><div class="sub ar">${esc(p.ar)}</div></div>
          <div class="ctl"><button class="cb ${sh ? 'on' : ''}" data-sh="${esc(p.c)}">✔ متوفر</button><div class="st"><button class="cb" data-od="-1" data-oc="${esc(p.c)}">−</button><b>${n}</b><button class="cb" data-od="1" data-oc="${esc(p.c)}">+</button></div></div></div>`;
        return `<div class="card pc" data-c="${esc(p.c)}">${X.img(p.c)}<div class="pcb"><div class="row"><b>${esc(p.c)}</b><span>${n ? `<span class="pill o">in cart ${n}</span> ` : ''}<span class="pill ${p.q <= 0 ? 'r' : 'g'}">${p.q <= 0 ? 'OUT' : p.q}</span></span></div><div class="sub">${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div></div></div>`; }).join('') + (L.length > lim ? `<button class="btn l" id="mr">Show more (${L.length - lim})</button>` : '') + (L.length === 0 ? '<div class="card">No items found.</div>' : '');
      const keep = fn => { const y = m.querySelector('div').scrollTop; fn(); drawList(); m.querySelector('div').scrollTop = y; };
      m.querySelectorAll('[data-sh]').forEach(b => b.onclick = e => { e.stopPropagation(); keep(() => toggleShelf(b.dataset.sh)); });
      m.querySelectorAll('[data-od]').forEach(b => b.onclick = e => { e.stopPropagation(); keep(() => setOrder(b.dataset.oc, ((cart.items.find(i => i.c === b.dataset.oc) || {}).q || 0) + +b.dataset.od)); });
      if (!R) { /* shelf / order buttons handle everything */ } else m.querySelectorAll('[data-c]').forEach(c => c.onclick = () => { const r = addItem(c.dataset.c, true); toast(r); if (!r.startsWith('Not found') && !r.startsWith('Only') && !r.includes('OUT OF')) { const y = m.querySelector('div').scrollTop; drawList(); m.querySelector('div').scrollTop = y; } });
      const mr = m.querySelector('#mr'); if (mr) mr.onclick = () => { lim += 40; const y = m.querySelector('div').scrollTop; drawList(); m.querySelector('div').scrollTop = y; };
    };
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelector('#si').oninput = e => { t = e.target.value.trim().toLowerCase(); lim = 40; drawList(); };
    drawCats(); drawList();
  }
  async function saveVisit(invNo) {   // one record per visit: what was on his shelf + the draft order (or the invoice number)
    try {
      const id = cart.visitId || doc(collection(X.db, 'shelf')).id, now = new Date();
      const rec = { cust: cart.cust.id, custEn: cart.cust.en, custAr: cart.cust.ar || '', shelf: cart.items.filter(i => i.shelf).map(i => i.c),
        draft: invNo ? [] : cart.items.filter(i => i.q > 0).map(i => ({ c: i.c, q: i.q })), invoiced: invNo || null, note: cart.note || '', branch: cart.branch || '', delivery: cart.delivery || '' };
      if (!cart.visitId) Object.assign(rec, { by: S.user.uid, byName: S.user.name, created: now.getTime(), date: iso(now) });
      const w = setDoc(doc(X.db, 'shelf', id), rec, { merge: true });
      // Offline, the write waits on the phone and syncs by itself later, so don't make the salesman wait for the server.
      const r = await Promise.race([w.then(() => 'ok'), new Promise(res => setTimeout(() => res('queued'), navigator.onLine ? 8000 : 600))]);
      w.catch(e => console.error(e)); cart.queued = (r === 'queued'); return true;
    } catch (e) { console.error(e); return false; }
  }
  async function saveOnly() {
    if (!cart.items.length) return;
    const b = $('#sv'); b.disabled = true; b.textContent = 'Saving…';
    brSave(cart.cust.id, cart.branch);
    if (!(await saveVisit(null))) { b.disabled = false; b.textContent = '💾 Save visit (shelf + draft order)'; return toast('Could not save the visit. Try again.'); }
    const n = cart.items.filter(i => i.shelf).length, o = cart.items.filter(i => i.q > 0).length, cu = cart.cust.en, qd = cart.queued;
    logEvent('visit', { cust: cart.cust.id, custEn: cu, shelf: n });
    S[F] = null; saveCart(kind);
    $('#app').innerHTML = top(qd ? 'Visit saved on phone ✓' : 'Visit saved ✓') + `<div class="wrap"><div class="card" style="text-align:center"><div style="font-size:44px">🧺</div><h2>${esc(cu)}</h2><div>${n} item(s) on his shelf</div>${qd ? '<div style="color:#b45309;margin-top:6px">No signal: it will upload automatically when the internet returns. Open the app once with signal. • سيتم الرفع تلقائيا عند عودة الإنترنت</div>' : ''}<div class="sub">${o ? o + ' item(s) saved as a draft order — open it from the Visits tab to make the invoice' : 'No order'}</div></div>
      <button class="btn g" id="nw">New visit</button><button class="btn l" id="hm">Home</button></div>`;
    $('#bk').onclick = () => go('home'); $('#nw').onclick = () => go('invoice'); $('#hm').onclick = () => go('home');
  }
  async function create() {
    if (!navigator.onLine) return alert('No internet. Invoices need a connection so stock stays correct.\nTap "Save visit" now and make the invoice later.\nلا يوجد اتصال: احفظ الزيارة وأنشئ الفاتورة لاحقا');
    const btn = $('#cr'); btn.disabled = true; btn.textContent = 'Creating…';
    const items = cart.items.filter(i => i.q > 0).map(i => ({ c: i.c, en: i.en, pk: i.pk, lav: lavOf(i.c), q: i.q, p: i.p, list: i.list }));
    const now = new Date(); let inv;
    try {
      await runTransaction(X.db, async tx => {
        const cref = doc(X.db, 'counters', 'invoice'), cs = await tx.get(cref), no = (cs.exists() ? cs.data().n : 0) + 1;
        const refs = items.map(i => doc(X.db, 'products', pid(i.c))), snaps = [];
        for (const r of refs) snaps.push(await tx.get(r));
        snaps.forEach((s, i) => { if (!s.exists() || s.data().q < items[i].q) throw new Error('STOCK:' + items[i].c); });
        snaps.forEach((s, i) => tx.update(refs[i], { q: s.data().q - items[i].q }));
        tx.set(cref, { n: no });
        inv = { no: 'INV-' + pad(no), cust: cart.cust.id, custEn: cart.cust.en, custAr: cart.cust.ar || '', disc: cart.cust.disc || 0, items,
          dm: 'end', subtotal: calc(items, cart.cust.disc).sub, discAmt: calc(items, cart.cust.disc).da, total: calc(items, cart.cust.disc).net, pieces: items.reduce((a, i) => a + i.q, 0), cartons: items.reduce((a, i) => a + i.q / cf(i), 0),
          by: S.user.uid, byName: S.user.name, created: now.getTime(), date: iso(now), delivery: cart.delivery || iso(now), note: cart.note || '', branch: (cart.branch || '').trim(), status: 'active', deliveryStatus: 'pending' };
        tx.set(doc(X.db, 'invoices', inv.no), inv);
      });
    } catch (e) {
      btn.disabled = false; btn.textContent = 'Create invoice';
      if (String(e.message).startsWith('STOCK:')) { await loadProducts(true); return alert('Not enough stock for ' + e.message.slice(6) + '. Stock was updated; please check quantities.'); }
      console.error(e); return toast('Could not create invoice. Try again.');
    }
    brSave(cart.cust.id, cart.branch); await saveVisit(inv.no);
    S.products = null; S.cart = null; saveCart('invoice'); logEvent('invoice', { ref: inv.no, cust: inv.cust, custEn: inv.custEn, total: inv.total }); emailInvoice(inv); done(inv);
  }
  async function createReturn() {
    const btn = $('#cr'); btn.disabled = true; btn.textContent = 'Saving…';
    const items = cart.items.map(i => ({ c: i.c, en: i.en, pk: i.pk, lav: lavOf(i.c), q: i.q, p: i.p, list: i.list, damaged: isDamaged(i.note) }));
    const now = new Date(), good = items.filter(i => !i.damaged); let ret;
    try {
      await runTransaction(X.db, async tx => {
        const cref = doc(X.db, 'counters', 'return'), cs = await tx.get(cref), no = (cs.exists() ? cs.data().n : 0) + 1;
        const refs = good.map(i => doc(X.db, 'products', pid(i.c))), snaps = [];
        for (const r of refs) snaps.push(await tx.get(r));
        snaps.forEach((s, i) => { if (s.exists()) tx.update(refs[i], { q: s.data().q + good[i].q }); });
        tx.set(cref, { n: no });
        ret = { no: 'RET-' + pad(no), cust: cart.cust.id, custEn: cart.cust.en, custAr: cart.cust.ar || '', disc: cart.cust.disc || 0, items,
          dm: 'end', subtotal: calc(items, cart.cust.disc).sub, discAmt: calc(items, cart.cust.disc).da, total: calc(items, cart.cust.disc).net, pieces: items.reduce((a, i) => a + i.q, 0), cartons: items.reduce((a, i) => a + i.q / cf(i), 0),
          restocked: good.reduce((a, i) => a + i.q, 0), damagedPcs: items.filter(i => i.damaged).reduce((a, i) => a + i.q, 0),
          by: S.user.uid, byName: S.user.name, created: now.getTime(), date: iso(now), note: cart.note || '', branch: (cart.branch || '').trim(), status: 'active' };
        tx.set(doc(X.db, 'returns', ret.no), ret);
      });
    } catch (e) { console.error(e); btn.disabled = false; btn.textContent = 'Create return'; return toast('Could not save the return. Try again.'); }
    brSave(cart.cust.id, cart.branch); S.products = null; S.rcart = null; saveCart('return'); logEvent('return', { ref: ret.no, cust: ret.cust, custEn: ret.custEn, total: ret.total }); emailReturn(ret); doneReturn(ret);
  }
  draw();
}
function done(inv) {
  const { top, esc, go, kwd, toast } = X;
  const fname = `${safe(nm(inv))} ${dmy(new Date(inv.created))}.xlsx`;
  $('#app').innerHTML = top('Invoice created ✓') + `<div class="wrap"><div class="card" style="text-align:center"><div style="font-size:44px">✅</div><h2>${esc(inv.no)}</h2>
    <div>${esc(nm(inv))}</div><div class="ar sub">${esc(inv.custAr)}</div><div style="font-size:24px;font-weight:800;margin-top:8px">${kwd(inv.total)} KWD</div>${inv.dm === 'end' && inv.disc ? `<div class="sub">${kwd(inv.subtotal)} − ${inv.disc}% discount ${kwd(inv.discAmt)}</div>` : ''}
    <div class="sub">${inv.pieces} pieces • ${inv.cartons.toFixed(2)} cartons</div></div>
    <button class="btn g" id="sh">📤 Share (WhatsApp / Email)</button><button class="btn" id="dl">⬇️ Download Excel</button><button class="btn l" id="nw">New visit</button><button class="btn l" id="hm">Home</button></div>`;
  $('#bk').onclick = () => go('home');
  const run = async mode => { try { await sendFile(await buildExcel(inv.items, fname, dOf(inv)), fname, mode); } catch (e) { console.error(e); toast('Could not build the Excel file'); } };
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
    <div>${esc(nm(i))}</div><div class="sub">${esc(i.date)} • by ${esc(i.byName)} • delivery ${esc(i.delivery)}</div></div>`).join('') : '<div class="card sub">No invoices yet.</div>';
  document.querySelectorAll('#rl .card[data-id]').forEach(c => c.onclick = () => openInvoice(L.find(i => i.id === c.dataset.id)));
}

async function recentVisits() {
  const { S, esc } = X;
  let L;
  try { L = (await getDocs(query(collection(X.db, 'shelf'), orderBy('created', 'desc'), limit(80)))).docs.map(d => ({ id: d.id, ...d.data() })); }
  catch (e) { return ($('#rl').innerHTML = '<div class="card">Could not load visits.</div>'); }
  if (S.user.role === 'sales') L = L.filter(v => v.by === S.user.uid);
  $('#rl').innerHTML = L.length ? L.map(v => `<div class="card" data-id="${v.id}"><div class="row"><b>${esc(nm(v))}</b><span class="pill ${v.invoiced ? 'g' : (v.draft || []).length ? 'o' : 'm'}">${v.invoiced ? esc(v.invoiced) : (v.draft || []).length ? 'draft order' : 'visit'}</span></div>
    <div class="sub">${esc(v.date)} • by ${esc(v.byName)} • ${(v.shelf || []).length} on shelf${(v.draft || []).length ? ' • ' + v.draft.length + ' ordered' : ''}</div></div>`).join('') : '<div class="card sub">No visits yet.</div>';
  document.querySelectorAll('#rl .card[data-id]').forEach(c => c.onclick = () => openVisit(L.find(v => v.id === c.dataset.id)));
}
async function openVisit(v) {
  const { S, esc } = X; await X.loadProducts();
  const P = c => S.products.find(p => p.c === c);
  const m = X.modal(`<h3>${esc(nm(v))}</h3><div class="sub">${esc(v.date)} • by ${esc(v.byName)}${v.invoiced ? ' • invoiced ' + esc(v.invoiced) : ''}</div>
    <b>On his shelf (${(v.shelf || []).length})</b>${(v.shelf || []).map(c => `<div class="sub">✔ ${esc(c)} ${esc((P(c) || {}).en || '')}</div>`).join('') || '<div class="sub">—</div>'}
    <b style="display:block;margin-top:8px">Draft order (${(v.draft || []).length})</b>${(v.draft || []).map(d => `<div class="sub">${esc(d.c)} × ${d.q}</div>`).join('') || '<div class="sub">—</div>'}
    ${v.invoiced ? '' : '<button class="btn g" id="op">Open to edit / make the invoice</button>'}<button class="btn l" id="cl">Close</button>`);
  m.querySelector('#cl').onclick = () => m.remove();
  const op = m.querySelector('#op'); if (!op) return;
  op.onclick = async () => {
    const cu = (await X.loadCustomers()).find(c => c.id === v.cust) || {}, disc = cu.disc || 0, items = [], miss = [];
    const add = (c, q, sh) => { let it = items.find(i => i.c === c); if (!it) { const p = P(c); if (!p) return miss.push(c); it = { c, en: p.en, ar: p.ar, pk: p.pk || 1, stock: p.q, list: p.p, p: p.p, q: 0, shelf: false, note: '' }; items.push(it); } if (q) it.q = Math.min(q, Math.max(0, it.stock)); if (sh) it.shelf = true; };
    (v.shelf || []).forEach(c => add(c, 0, true)); (v.draft || []).forEach(d => add(d.c, d.q, false));
    X.S[KEYS.invoice[0]] = { cust: { id: v.cust, en: v.custEn, ar: v.custAr || '', disc }, items, delivery: v.delivery || iso(new Date(Date.now() + 864e5)), note: v.note || '', branch: v.branch || '', visitId: v.id };
    saveCart('invoice'); m.remove(); if (miss.length) X.toast('Not found: ' + miss.join(', ')); cartView('invoice');
  };
}

function openInvoice(inv) {
  const { S, esc, kwd, toast } = X;
  const fname = `${safe(nm(inv))} ${dmy(new Date(inv.created))}.xlsx`, admin = S.user.role === 'admin';
  const m = X.modal(`<h3>${esc(inv.no)} ${inv.status === 'cancelled' ? '<span class="pill r">CANCELLED</span>' : ''}</h3><div>${esc(nm(inv))}</div><div class="ar sub">${esc(inv.custAr)}</div>
    <div class="sub">${esc(inv.date)} • by ${esc(inv.byName)} • delivery ${esc(inv.delivery)}</div>
    ${inv.items.map(i => `<div class="row"><span>${esc(i.c)} × ${i.q}</span><span>${kwd(i.q * i.p)}</span></div>`).join('')}
    ${inv.dm === 'end' && inv.disc ? `<div class="row" style="margin-top:8px"><span>Subtotal</span><span>${kwd(inv.subtotal)}</span></div><div class="row"><span>Discount ${inv.disc}%</span><span>-${kwd(inv.discAmt)}</span></div>` : ''}<div class="row" style="margin-top:8px"><b>Total</b><b>${kwd(inv.total)} KWD</b></div>
    <button class="btn g" id="sh">📤 Share Excel</button><button class="btn" id="dl">⬇️ Download Excel</button>
    ${admin && inv.status === 'active' ? '<button class="btn r" id="cx">Cancel this invoice (stock returns)</button>' : ''}<button class="btn l" id="cl">Close</button>`);
  const run = async mode => { try { await sendFile(await buildExcel(inv.items, fname, dOf(inv)), fname, mode); } catch (e) { toast('Could not build the Excel file'); } };
  m.querySelector('#sh').onclick = () => run('share'); m.querySelector('#dl').onclick = () => run('download'); m.querySelector('#cl').onclick = () => m.remove();
  const cx = m.querySelector('#cx');
  if (cx) cx.onclick = async () => {
    if (!confirm(`Cancel ${inv.no}? The items go back into stock.`)) return;
    try {
      await X.loadProducts();
      await runTransaction(X.db, async tx => {
        const ir = doc(X.db, 'invoices', inv.id), is = await tx.get(ir); if (is.data().status !== 'active') throw new Error('done');
        const refs = inv.items.map(i => doc(X.db, 'products', pid(i.c))), snaps = [];
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
  const fname = `${safe(nm(ret))} RETURN ${dmy(new Date(ret.created))}.xlsx`;
  $('#app').innerHTML = top('Return saved ✓') + `<div class="wrap"><div class="card" style="text-align:center"><div style="font-size:44px">↩️</div><h2>${esc(ret.no)}</h2>
    <div>${esc(nm(ret))}</div><div class="ar sub">${esc(ret.custAr)}</div><div style="font-size:24px;font-weight:800;margin-top:8px">${kwd(ret.total)} KWD credit</div>
    <div class="sub">${ret.pieces} pieces • back in stock: ${ret.restocked} • damaged: ${ret.damagedPcs}</div></div>
    <button class="btn g" id="sh">📤 Share (WhatsApp / Email)</button><button class="btn" id="dl">⬇️ Download Excel</button><button class="btn l" id="nw">New return</button><button class="btn l" id="hm">Home</button></div>`;
  $('#bk').onclick = () => go('home');
  const run = async mode => { try { await sendFile(await buildExcel(ret.items, fname, dOf(ret)), fname, mode); } catch (e) { console.error(e); toast('Could not build the Excel file'); } };
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
    <div>${esc(nm(i))}</div><div class="sub">${esc(i.date)} • by ${esc(i.byName)} • damaged ${i.damagedPcs}</div></div>`).join('') : '<div class="card sub">No returns yet.</div>';
  document.querySelectorAll('#rl .card[data-id]').forEach(c => c.onclick = () => {
    const r = L.find(i => i.id === c.dataset.id), fname = `${safe(nm(r))} RETURN ${dmy(new Date(r.created))}.xlsx`;
    const m = X.modal(`<h3>${esc(r.no)}</h3><div>${esc(nm(r))}</div><div class="sub">${esc(r.date)} • by ${esc(r.byName)}</div>
      ${r.items.map(i => `<div class="row"><span>${esc(i.c)} × ${i.q} ${i.damaged ? '<span class="pill r">damaged</span>' : ''}</span><span>${kwd(i.q * i.p)}</span></div>`).join('')}
      ${r.dm === 'end' && r.disc ? `<div class="row" style="margin-top:8px"><span>Subtotal</span><span>${kwd(r.subtotal)}</span></div><div class="row"><span>Discount ${r.disc}%</span><span>-${kwd(r.discAmt)}</span></div>` : ''}<div class="row" style="margin-top:8px"><b>Credit</b><b>${kwd(r.total)} KWD</b></div>
      <button class="btn g" id="sh">📤 Share Excel</button><button class="btn" id="dl">⬇️ Download Excel</button><button class="btn l" id="cl">Close</button>`);
    const run = async mode => { try { await sendFile(await buildExcel(r.items, fname, dOf(r)), fname, mode); } catch (e) { toast('Could not build the Excel file'); } };
    m.querySelector('#sh').onclick = () => run('share'); m.querySelector('#dl').onclick = () => run('download'); m.querySelector('#cl').onclick = () => m.remove();
  });
}
