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
export async function sendFile(blob, filename, mode) {
  const f = new File([blob], filename, { type: blob.type });
  if (mode === 'share' && navigator.canShare && navigator.canShare({ files: [f] })) {
    try { await navigator.share({ files: [f], title: filename }); return; } catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ---------- Barcode scanner (camera) ---------- */
export function scanner(onCode) {
  const m = X.modal(`<h3>Scan barcode</h3><video id="vd" playsinline muted style="width:100%;border-radius:12px;background:#000;max-height:50vh"></video>
    <div class="sub" id="sm">Starting camera…</div><button class="btn l" id="cl">Close</button>`);
  let reader, stopped = false, last = '', lastT = 0, pend = '', pendN = 0;
  const stop = () => { stopped = true; try { reader && reader.reset(); } catch (e) {} m.remove(); };
  m.querySelector('#cl').onclick = stop;
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

/* ---------- New invoice ---------- */
const CART_KEY = 'kw_cart';
function saveCart() { try { localStorage.setItem(CART_KEY, JSON.stringify(X.S.cart || null)); } catch (e) {} }
function restoreCart() { if (X.S.cart) return; try { X.S.cart = JSON.parse(localStorage.getItem(CART_KEY) || 'null'); } catch (e) {} }

export async function invoiceHome() {
  const { S, top, esc, go, kwd } = X;
  restoreCart();
  if (S.cart && S.cart.items) return cartView();
  customerPick();
}

async function customerPick() {
  const { S, top, esc, go, kwd, loadCustomers } = X;
  $('#app').innerHTML = top('New Invoice • فاتورة جديدة') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  const C = await loadCustomers(); let q = '', tab = 'new';
  const draw = () => {
    const t = q.trim().toLowerCase();
    const L = C.filter(c => (c.active !== false || t) && (!t || c.id.includes(t) || (c.en || '').toLowerCase().includes(t) || (c.ar || '').includes(t))).sort((a, b) => (a.en || '').localeCompare(b.en || ''));
    $('#app').innerHTML = top('New Invoice • فاتورة جديدة') + `<div class="wrap">
      <div class="tabs"><button data-t="new" class="${tab === 'new' ? 'on' : ''}">New</button><button data-t="recent" class="${tab === 'recent' ? 'on' : ''}">Recent invoices</button></div>
      ${tab === 'new' ? `<div class="sub">Choose the customer</div><input id="q" placeholder="Search customer…" value="${esc(q)}">
      ${L.map(c => `<div class="card" data-id="${c.id}"><b>${esc(c.en)}</b><div class="ar sub">${esc(c.ar)}</div>${c.disc ? `<span class="pill g">${c.disc}% discount</span>` : ''}</div>`).join('')}` : '<div id="rl"><div class="sub">Loading…</div></div>'}</div>`;
    $('#bk').onclick = () => go('home');
    document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; draw(); });
    if (tab === 'new') {
      const qi = $('#q'); qi.oninput = e => { q = e.target.value; draw(); const n = $('#q'); n.focus(); n.setSelectionRange(q.length, q.length); };
      document.querySelectorAll('.card[data-id]').forEach(c => c.onclick = () => {
        const cu = C.find(x => x.id === c.dataset.id);
        S.cart = { cust: { id: cu.id, en: cu.en, ar: cu.ar, disc: cu.disc || 0 }, items: [], delivery: iso(new Date(Date.now() + 864e5)), note: '' }; saveCart(); cartView();
      });
    } else recentList();
  };
  draw();
}

async function cartView() {
  const { S, top, esc, go, kwd, toast, loadProducts } = X;
  const cart = S.cart; await loadProducts();
  if (!S.barcodes) { S.barcodes = {}; try { (await getDocs(collection(X.db, 'barcodes'))).docs.forEach(d => S.barcodes[d.id] = d.data().c); } catch (e) {} }
  const bkey = c => String(c).trim().replace(/\//g, '_');
  const P = S.products, by = id => { const k = String(id).trim(); let p = P.find(p => p.id === k || p.c.toLowerCase() === k.toLowerCase()); if (!p && S.barcodes[bkey(k)]) p = P.find(p => p.c === S.barcodes[bkey(k)]); return p; };
  const total = () => r3(cart.items.reduce((a, i) => a + i.q * i.p, 0));
  const draw = () => {
    const pcs = cart.items.reduce((a, i) => a + i.q, 0), ctn = cart.items.reduce((a, i) => a + i.q / i.pk, 0);
    $('#app').innerHTML = top('Invoice • ' + esc(cart.cust.en)) + `<div class="wrap">
      ${cart.cust.disc ? `<div class="pill g" style="margin-bottom:8px">Customer discount ${cart.cust.disc}%</div>` : ''}
      <div class="grid" style="margin-bottom:10px"><button class="btn" id="sc">📷 Scan</button><button class="btn l" id="se">🔎 Search item</button></div>
      ${cart.items.length === 0 ? '<div class="card sub">No items yet. Scan a barcode or search for an item.</div>' : ''}
      ${cart.items.map((it, i) => `<div class="card"><div class="row"><b>${esc(it.c)}</b><button class="pill r" data-rm="${i}" style="border:0">Remove</button></div>
        <div class="sub">${esc(it.en)}</div><div class="ar sub">${esc(it.ar)}</div>
        <div class="row" style="margin-top:6px"><div style="flex:1"><label>Pieces (stock ${it.stock})</label><input type="number" inputmode="numeric" data-q="${i}" value="${it.q}"></div>
        <button class="btn l" style="width:auto;margin:20px 0 0" data-ctn="${i}">+1 carton<br><span class="sub">(${it.pk} pcs)</span></button></div>
        <div class="sub">${kwd(it.p)} KWD each • ${(it.q / it.pk).toFixed(2)} cartons • <b>${kwd(it.q * it.p)}</b></div></div>`).join('')}
      <div class="card"><div class="row"><span>Pieces</span><b>${pcs}</b></div><div class="row"><span>Cartons</span><b>${ctn.toFixed(2)}</b></div>
        <div class="row"><span>Total (KWD)</span><b style="font-size:22px">${kwd(total())}</b></div></div>
      <label>Delivery date • تاريخ التوصيل</label><input type="date" id="dd" value="${cart.delivery}">
      <label>Note (optional)</label><input id="nt" value="${esc(cart.note)}">
      <button class="btn g" id="cr" ${cart.items.length ? '' : 'disabled style="opacity:.5"'}>Create invoice</button>
      <button class="btn l" id="cn">Cancel / start over</button></div>`;
    $('#bk').onclick = () => { go('home'); };
    document.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { cart.items.splice(+b.dataset.rm, 1); saveCart(); draw(); });
    document.querySelectorAll('[data-q]').forEach(inp => inp.onchange = () => {
      const it = cart.items[+inp.dataset.q]; let v = Math.max(1, Math.floor(+inp.value || 1));
      if (v > it.stock) { v = it.stock; toast(`Only ${it.stock} in stock`); } it.q = v; saveCart(); draw();
    });
    document.querySelectorAll('[data-ctn]').forEach(b => b.onclick = () => {
      const it = cart.items[+b.dataset.ctn]; if (it.q + it.pk > it.stock) { toast(`Only ${it.stock} in stock`); it.q = it.stock; } else it.q += it.pk; saveCart(); draw();
    });
    $('#dd').onchange = e => { cart.delivery = e.target.value; saveCart(); };
    $('#nt').onchange = e => { cart.note = e.target.value; saveCart(); };
    $('#cn').onclick = () => { if (!cart.items.length || confirm('Throw away this invoice draft?')) { S.cart = null; saveCart(); go('invoice'); } };
    $('#sc').onclick = () => scanner(code => {
      if (!by(code)) { setTimeout(() => linkBarcode(code), 50); return false; }
      toast(addItem(code)); return true;
    });
    $('#se').onclick = searchItem;
    $('#cr').onclick = create;
  };
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
        try { await setDoc(doc(X.db, 'barcodes', bkey(code)), { c: c.dataset.c, by: S.user.name, at: Date.now() }); S.barcodes[bkey(code)] = c.dataset.c; }
        catch (e) { return toast('Could not save the barcode'); }
        m.remove(); toast(addItem(c.dataset.c) + ' • barcode saved');
      });
    };
  }
  function addItem(code) {
    const p = by(code); if (!p) return `Not found: ${code}`;
    if (p.q <= 0) return `${p.c} is OUT OF STOCK`;
    const ex = cart.items.find(i => i.c === p.c);
    if (ex) { if (ex.q >= ex.stock) return `Only ${ex.stock} in stock`; ex.q += 1; }
    else cart.items.push({ c: p.c, en: p.en, ar: p.ar, pk: p.pk || 1, stock: p.q, list: p.p, p: r3(p.p * (1 - (cart.cust.disc || 0) / 100)), q: 1 });
    saveCart(); draw(); return `Added ${p.c}`;
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
    S.products = null; S.cart = null; saveCart(); done(inv);
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
      X.S.products = null; m.remove(); toast('Invoice cancelled, stock returned'); X.go('invoice');
    } catch (e) { toast('Could not cancel'); }
  };
}
