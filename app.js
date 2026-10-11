import { firebaseConfig } from './firebase-config.js';
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import { getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, updatePassword } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { getFirestore, initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, getDoc, setDoc, getDocs, collection, writeBatch, updateDoc, deleteDoc } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

const $app = document.getElementById('app');
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const kwd = n => Number(n || 0).toLocaleString('en-US', {minimumFractionDigits: 3, maximumFractionDigits: 3});
const toast = m => { const t = document.getElementById('toast'); t.textContent = m; t.className = 'on'; setTimeout(() => t.className = '', 2600); };
if ('serviceWorker' in navigator) {
  const had = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register('sw.js').then(r => r.update()).catch(() => {});
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (had) location.reload(); });
}

if (!firebaseConfig) {
  $app.innerHTML = '<div class="wrap"><div class="card"><b>Almost ready</b><p>The app is built but not yet connected to your database. Waiting for the Firebase settings.</p></div></div>';
  throw new Error('no config');
}

import { initInvoices, invoiceHome, returnHome, buildStockExcel, sendFile } from './invoices.js';
import { initStatement, statementHome, chequesHome } from './statement.js';
import { initActivity, getLoc } from './activity.js';
import { initCalendar, calendarHome } from './calendar.js';
import { initReport, reportHome } from './report.js';
import { initBackup, backupHome } from './backup.js';
import { initShelf, shelfHome } from './shelf.js';
const fb = initializeApp(firebaseConfig);
const auth = getAuth(fb);
// Keep a copy of the data on the phone so the app opens and visits can be saved without signal (they sync later).
let db; try { db = initializeFirestore(fb, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) }); } catch (e) { db = getFirestore(fb); }
(function offlineBanner() {
  const b = document.createElement('div'); b.id = 'offb';
  b.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:99;background:#b45309;color:#fff;text-align:center;padding:6px 10px;font-size:13px;display:none';
  b.textContent = 'No internet • You can still do Shelf Check visits; they sync when the signal returns. Invoices need internet. • لا يوجد اتصال';
  document.addEventListener('DOMContentLoaded', () => document.body.appendChild(b));
  const u = () => { b.style.display = navigator.onLine ? 'none' : 'block'; };
  addEventListener('online', u); addEventListener('offline', u); document.addEventListener('DOMContentLoaded', u);
})();
const mail = u => u.trim().toLowerCase().replace(/[^a-z0-9]/g, '') + '@kitchenware.app';
const pw = pin => pin + '-kw';

const S = { user: null, products: null, customers: null, lowLimit: 5, view: 'home' };
const LOW = () => S.lowLimit;
let IMGMAP = {}, HIDDEN = new Set(), HIDEONLY = new Set(), RENAMES = {}, ARFIX = {}, FDESC = {};
const jget = u => fetch(u, { cache: 'no-store' }).then(r => r.json()).catch(() => null);
const READY = Promise.all([jget('./imgmap.json').then(m => { if (m) IMGMAP = m; }), jget('./hidden.json').then(m => { if (m) HIDDEN = new Set(m); }), jget('./hiddenonly.json').then(m => { if (m) HIDEONLY = new Set(m); }), jget('./renames.json').then(m => { if (m) RENAMES = m; }), jget('./arfix.json').then(m => { if (m) ARFIX = m; }), jget('./flaskdesc.json').then(m => { if (m) FDESC = m; })]);
const isHidden = p => p.hidden === undefined ? (HIDDEN.has(p.c) || HIDEONLY.has(p.c)) : !!p.hidden;
const toDelete = p => isHidden(p) && p.hidden === undefined && HIDDEN.has(p.c);
let CATMAP = {}; const READY2 = jget('./cats.json').then(m => { if (m) CATMAP = m; });
const CATLIST = [['lav', 'LAV'], ['flask', 'Flask • مطارات'], ['hascevher', 'Hascevher'], ['porcelain', 'Porcelain • بورسلين'], ['tea', 'Tea Set • طقم شاي'], ['tray', 'Tray • صواني'], ['others', 'Others • أخرى']];
let SUBMAP = {}; jget('./subcats.json').then(m => { if (m) SUBMAP = m; });
const SUBS = { flask: [['sun', 'SUN'], ['shimizu', 'SHIMIZU'], ['fg', 'FG'], ['other', 'Others']],
  hascevher: [['pots', 'قدور'], ['fry', 'مقالي'], ['other', 'Others']],
  lav: [['tumbler', 'كاسات'], ['stem', 'كاسات بقاعدة'], ['tea', 'استكانة / صحن'], ['mug', 'ماج'], ['dolka', 'دولكة'], ['other', 'Others']] };
const subOf = p => p.sub || SUBMAP[p.c] || 'other';
const subRank = p => { const L = SUBS[catOf(p)]; return L ? Math.max(0, L.findIndex(s => s[0] === subOf(p))) : 0; };
const catOf = p => p.cat || CATMAP[p.c] || 'others';
const pimg = (c, big) => IMGMAP[c] ? `<img class="pt${big ? ' big' : ''}" loading="lazy" src="./img/${IMGMAP[c]}" data-full="./img/${IMGMAP[c]}" alt="">` : '';
document.addEventListener('click', e => { const t = e.target; if (t.classList && t.classList.contains('pt')) { e.stopPropagation(); e.preventDefault(); const m = modal(`<img src="${t.dataset.full}" style="width:100%;max-height:70vh;object-fit:contain;background:#fff;border-radius:12px">${/\/fx_/.test(t.dataset.full) ? '<div class="sub" style="text-align:center;margin-top:6px">Edited picture — colour changed digitally, not a real photo</div>' : ''}<button class="btn l" id="zc">Close</button>`); m.querySelector('#zc').onclick = () => m.remove(); } }, true);
const CTX = { SUBS, subOf: p => subOf(p), subRank: p => subRank(p), isHidden: p => isHidden(p), catOf: p => catOf(p), CATLIST, img: (c, b) => pimg(c, b), db, S, top: (...a) => top(...a), esc, go: (...a) => go(...a), kwd, toast, modal: h => modal(h), loadProducts: f => loadProducts(f), loadCustomers: f => loadCustomers(f), loadBarcodes: f => loadBarcodes(f), editCustomer: (c, f) => editCustomer(c, f) };
initInvoices(CTX); initStatement(CTX); initActivity(CTX); initCalendar(CTX); initReport(CTX); initBackup(CTX); initShelf(CTX);

const ROLE_HOME = {
  admin: ['invoice', 'return', 'stock', 'statement', 'shelf', 'report', 'calendar'],
  sales: ['invoice', 'return', 'stock', 'statement', 'shelf', 'calendar'],
  warehouse: ['calendar']
};
const BTN = {
  invoice: ['🧺', 'Shelf Check / Invoice', 'جرد الرف / فاتورة'], shelf: ['📈', 'Shelf Report', 'تقرير الرف'], return: ['↩️', 'Return', 'مرتجع'], stock: ['📦', 'Stock', 'المخزون'],
  statement: ['👤', 'Customer Statement', 'كشف حساب العميل'], report: ['📊', 'Daily Report', 'التقرير اليومي'], calendar: ['📅', 'Calendar', 'التقويم']
};

onAuthStateChanged(auth, async u => {
  if (!u) { S.user = null; return showLogin(); }
  try {
    const d = await getDoc(doc(db, 'users', u.uid));
    if (!d.exists() || d.data().active === false) { await signOut(auth); toast('Account not active'); return; }
    S.user = { uid: u.uid, ...d.data() };
    go('home');
  } catch (e) { if (S.setupBusy) return; console.error(e); showLogin('Cannot load account'); }
});

async function showLogin(msg) {
  let needSetup = false;
  try { needSetup = !(await getDoc(doc(db, 'config', 'setup'))).exists(); } catch (e) {}
  let pin = '';
  const draw = () => {
    $app.innerHTML = `<div class="login">
      <div style="font-size:56px">🍳</div><h1>${needSetup ? 'First-time setup' : 'Welcome'}</h1>
      <div class="sub">${needSetup ? 'Create the first admin account (you)' : 'مرحباً'}</div><br>
      <input id="un" placeholder="Username (letters/numbers)" autocapitalize="none" autocomplete="off" value="${esc(window._un || '')}">
      ${needSetup ? '<input id="dn" placeholder="Your full name">' : ''}
      <div class="sub">PIN — 4 digits</div>
      <div class="pin">${[0,1,2,3].map(i => `<i class="${i < pin.length ? 'f' : ''}"></i>`).join('')}</div>
      <div class="err">${esc(msg || '')}</div>
      <div class="pad">${[1,2,3,4,5,6,7,8,9,'⌫',0,'OK'].map(k => `<button data-k="${k}">${k}</button>`).join('')}</div></div>`;
    $app.querySelectorAll('.pad button').forEach(b => b.onclick = async () => {
      const k = b.dataset.k; const un = document.getElementById('un'); window._un = un.value;
      if (k === '⌫') pin = pin.slice(0, -1); else if (k === 'OK') return submit(); else if (pin.length < 4) pin += k;
      if (pin.length === 4 && k !== '⌫') { const dn = document.getElementById('dn'); window._dn = dn ? dn.value : ''; }
      const dnv = document.getElementById('dn')?.value; draw(); if (dnv) document.getElementById('dn').value = dnv;
      if (pin.length === 4 && k !== '⌫') submit();
    });
  };
  const submit = async () => {
    const un = window._un || document.getElementById('un')?.value || '';
    const dn = document.getElementById('dn')?.value || window._dn || '';
    if (!un.trim() || pin.length !== 4) { msg = 'Enter username and 4-digit PIN'; return draw(); }
    try {
      if (needSetup) {
        S.setupBusy = true;
        const c = await createUserWithEmailAndPassword(auth, mail(un), pw(pin));
        const b = writeBatch(db);
        b.set(doc(db, 'users', c.user.uid), { username: un.trim().toLowerCase(), name: dn || un, role: 'admin', active: true, created: Date.now() });
        b.set(doc(db, 'config', 'setup'), { done: true, at: Date.now(), lowLimit: 5 });
        await b.commit(); S.setupBusy = false;
        S.user = { uid: c.user.uid, name: dn || un, role: 'admin', active: true }; go('home');
      } else await signInWithEmailAndPassword(auth, mail(un), pw(pin));
    } catch (e) { S.setupBusy = false; pin = ''; msg = 'Wrong username or PIN'; draw(); }
  };
  draw();
}

function top(title, back = true) {
  return `<div class="top">${back ? '<button id="bk">←</button>' : ''}<h1>${title}</h1>${S.view === 'home' ? '<button id="lo">Logout</button>' : ''}</div>`;
}
function go(v, arg) {
  S.view = v; S.arg = arg;
  const fn = { home, stock, statement: statementHome, cheques: chequesHome, calendar: calendarHome, report: reportHome, backup: backupHome, shelf: shelfHome, staff, invoice: invoiceHome, return: returnHome, comingsoon }[v] || comingsoon;
  fn();
  const bk = document.getElementById('bk'); if (bk) bk.onclick = () => go('home');
  const lo = document.getElementById('lo'); if (lo) lo.onclick = () => signOut(auth);
  window.scrollTo(0, 0);
}

function home() {
  const r = S.user.role, items = ROLE_HOME[r] || [];
  const adminExtra = r === 'admin' ? `<div class="grid" style="margin-top:12px">
    <button class="big small" data-go="cheques"><span class="ic">💳</span>Cheques<span class="ar">الشيكات</span></button>
    <button class="big small" data-go="staff"><span class="ic">👥</span>Staff<span class="ar">الموظفون</span></button>
    <button class="big small" data-go="backup"><span class="ic">💾</span>Backup<span class="ar">نسخة احتياطية</span></button></div>` : '';
  $app.innerHTML = top(`Hello, ${esc(S.user.name)}`, false).replace('<h1>', '<h1 style="font-size:18px">') + `<div class="wrap"><div class="grid">${items.map(k =>
    `<button class="big" data-go="${k}"><span class="ic">${BTN[k][0]}</span>${BTN[k][1]}<span class="ar">${BTN[k][2]}</span></button>`).join('')}</div>${adminExtra}
    <p class="sub" style="text-align:center;margin-top:24px">${esc(S.user.role)} • ${esc(S.user.username || '')}</p></div>`;
  $app.querySelectorAll('[data-go]').forEach(b => b.onclick = () => go(b.dataset.go));
  if (r !== 'admin' && !S.locAsked) { S.locAsked = true; getLoc(); } // asks the employee for location permission once
  const lo = document.getElementById('lo'); if (lo) lo.onclick = () => signOut(auth);
}

function comingsoon() {
  $app.innerHTML = top('Coming next') + '<div class="wrap"><div class="card">This part is being built in the next steps.</div></div>';
}

async function loadProducts(force) {
  if (S.products && !force) return S.products;
  const q = await getDocs(collection(db, 'products'));
  await READY; await READY2;
  S.products = q.docs.map(d => { const x = { id: d.id, ...d.data() }; if (RENAMES[x.c]) x.c = RENAMES[x.c]; if (ARFIX[x.c] && !x.arEdited) x.ar = ARFIX[x.c]; const f = FDESC[x.c]; if (f) { if (!x.enEdited) x.en = f.en; if (!x.arEdited) x.ar = f.ar; } return x; });
  return S.products;
}
async function loadBarcodes(force) {
  if (S.barcodes && !force) return S.barcodes;
  S.barcodes = {}; S.barcodeCodes = {};
  try { (await getDocs(collection(db, 'barcodes'))).docs.forEach(d => { S.barcodes[d.id] = d.data().c; S.barcodeCodes[d.id] = d.data().code || d.id; }); } catch (e) {}
  return S.barcodes;
}
function invBarcodes() { // item number -> list of learned barcodes
  const m = {}; Object.keys(S.barcodes || {}).forEach(id => { const c = S.barcodes[id]; (m[c] ||= []).push({ id, code: (S.barcodeCodes || {})[id] || id }); }); return m;
}
async function loadCustomers(force) {
  if (S.customers && !force) return S.customers;
  const q = await getDocs(collection(db, 'customers'));
  S.customers = q.docs.map(d => ({ id: d.id, ...d.data() }));
  return S.customers;
}
async function loadConfig() { try { const d = await getDoc(doc(db, 'config', 'setup')); S.lowLimit = d.data()?.lowLimit ?? 5; } catch (e) {} }

async function stock() {
  const admin = S.user.role === 'admin';
  $app.innerHTML = top('Stock • المخزون') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  document.getElementById('bk').onclick = () => go('home');
  await loadConfig(); let P = await loadProducts(); await loadBarcodes(); let hiddenView = false; let IB = invBarcodes();
  let filter = 'all', q = '', cat = 'all', sub = 'all';
  const draw = () => {
    const VP = P.filter(p => !isHidden(p) && p.q > 0), nh = P.filter(isHidden).length, nz = P.filter(p => !isHidden(p) && p.q <= 0).length, low = VP.filter(p => p.q > 0 && p.q <= LOW()).length, zero = nz;
    const t = q.trim().toLowerCase();
    let L = P.filter(p => (hiddenView ? isHidden(p) : !isHidden(p)) && (cat === 'all' || catOf(p) === cat) && (sub === 'all' || subOf(p) === sub) && (filter === 'all' ? (p.q > 0 || hiddenView || !!t) : (filter === 'low' && p.q > 0 && p.q <= LOW()) || (filter === 'zero' && p.q <= 0)) &&
      (!t || p.c.toLowerCase().includes(t) || (p.en || '').toLowerCase().includes(t) || (p.ar || '').includes(t) || (IB[p.c] || []).some(b => b.code.toLowerCase().includes(t))));
    if (cat !== 'all' && SUBS[cat]) L = L.slice().sort((x, y) => subRank(x) - subRank(y) || x.c.localeCompare(y.c));
    const shown = L.slice(0, 80);
    $app.innerHTML = top('Stock • المخزون') + `<div class="wrap">
      <input id="q" placeholder="Search item number, name or barcode…" value="${esc(q)}">
      ${admin ? '<div class="row"><button class="btn l" id="xd">⬇ Export Excel</button><button class="btn l" id="xs">Share</button></div><button class="btn l" id="up">⬆ Update stock from file</button>' : ''}
 <div class="tabs cats"><button data-k="all" class="${cat === 'all' ? 'on' : ''}">All</button>${CATLIST.map(([k, n]) => `<button data-k="${k}" class="${cat === k ? 'on' : ''}">${n} ${P.filter(p => !isHidden(p) && p.q > 0 && catOf(p) === k).length}</button>`).join('')}</div>
      ${cat !== 'all' && SUBS[cat] ? `<div class="tabs cats sub"><button data-s="all" class="${sub === 'all' ? 'on' : ''}">All</button>${SUBS[cat].map(([k, n]) => `<button data-s="${k}" class="${sub === k ? 'on' : ''}">${n} ${P.filter(p => !isHidden(p) && p.q > 0 && catOf(p) === cat && subOf(p) === k).length}</button>`).join('')}</div>` : ''}
      <div class="tabs"><button data-f="all" class="${filter === 'all' && !hiddenView ? 'on' : ''}">All ${VP.length}</button>
      <button data-f="low" class="${filter === 'low' ? 'on' : ''}">Low ${low}</button>
      <button data-f="zero" class="${filter === 'zero' ? 'on' : ''}">Out (hidden) ${zero}</button>${admin && nh ? `<button data-f="hid" class="${hiddenView ? 'on' : ''}">Hidden ${nh}</button>` : ''}</div>
      ${hiddenView && admin && P.some(toDelete) ? `<button class="btn" id="dh" style="background:#c0392b;color:#fff">🗑 Delete the ${P.filter(toDelete).length} items marked for deletion</button>` : ''}
      ${shown.map(p => `<div class="card pc" data-id="${esc(p.id)}">${pimg(p.c)}<div class="pcb"><div class="row"><b>${esc(p.c)}</b>
        <span class="pill ${p.q <= 0 ? 'r' : p.q <= LOW() ? 'o' : 'g'}">${p.q <= 0 ? 'OUT' : p.q} ${p.q > 0 ? 'pcs' : ''}</span></div>
        <div>${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div>
        ${(IB[p.c] || []).length ? `<div class="sub">▮▯ ${(IB[p.c]).map(b => esc(b.code)).join(' • ')}</div>` : ''}<div class="sub">Packing ${p.pk} • ${kwd(p.p)} KWD${p.pk && p.q > 0 ? ` • ${(p.q / p.pk).toFixed(1)} cartons` : ''}</div></div></div>`).join('')}
      ${L.length > 80 ? `<div class="sub" style="text-align:center">Showing 80 of ${L.length}. Type to narrow down.</div>` : ''}
      ${L.length === 0 ? '<div class="card">No items found.</div>' : ''}</div>`;
    document.getElementById('bk').onclick = () => go('home');
    const qi = document.getElementById('q'); qi.oninput = e => { q = e.target.value; draw(); const n = document.getElementById('q'); n.focus(); n.setSelectionRange(q.length, q.length); };
    $app.querySelectorAll('.tabs:not(.cats) button').forEach(b => b.onclick = () => { if (b.dataset.f === 'hid') { hiddenView = true; filter = 'all'; } else { hiddenView = false; filter = b.dataset.f; } draw(); });
    $app.querySelectorAll('.cats:not(.sub) button').forEach(b => b.onclick = () => { cat = b.dataset.k; sub = 'all'; draw(); });
    $app.querySelectorAll('.cats.sub button').forEach(b => b.onclick = () => { sub = b.dataset.s; draw(); });
    const dh = document.getElementById('dh');
    if (dh) dh.onclick = async () => {
      const H = P.filter(toDelete), pcs = H.reduce((s, p) => s + Math.max(0, p.q), 0);
      if (!confirm(`Permanently delete ${H.length} hidden items (${pcs} pieces in stock)?\n\nA backup Excel of these items will download first. This cannot be undone.`)) return;
      dh.disabled = true; dh.textContent = 'Saving backup…';
      try {
        const d = new Date(), n = `Deleted items backup ${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}.xlsx`;
        await sendFile(await buildStockExcel(H, {}), n, 'download');
      } catch (e) { if (!confirm('Backup could not be saved. Delete anyway?')) { draw(); return; } }
      dh.textContent = 'Deleting…';
      try {
        const ids = new Set(H.map(p => p.id)), codes = new Set(H.map(p => p.c));
        const refs = H.map(p => doc(db, 'products', p.id));
        await loadBarcodes();
        Object.keys(S.barcodes).forEach(b => { if (codes.has(S.barcodes[b])) refs.push(doc(db, 'barcodes', b)); });
        for (let i = 0; i < refs.length; i += 400) { const b = writeBatch(db); refs.slice(i, i + 400).forEach(r => b.delete(r)); await b.commit(); }
        P = S.products = P.filter(p => !ids.has(p.id)); S.barcodes = null; await loadBarcodes(true); IB = invBarcodes();
        hiddenView = false; toast(`${H.length} items deleted`); draw();
      } catch (e) { console.error(e); toast('Delete failed — nothing more was removed'); draw(); }
    };
    if (admin) {
      const run = async mode => { try { const d = new Date(), n = `Stock ${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}.xlsx`; const by = {}; Object.keys(IB).forEach(k => by[k] = IB[k].map(b => b.code)); await sendFile(await buildStockExcel(P, by, p => isHidden(p) ? (toDelete(p) ? 'MARKED FOR DELETION — remove from your files' : 'HIDDEN — remove from your files') : ''), n, mode); } catch (e) { console.error(e); toast('Could not build the Excel file'); } };
      document.getElementById('up').onclick = () => stockUpdate(() => { P = S.products; IB = invBarcodes(); draw(); });
      document.getElementById('xd').onclick = () => run('download'); document.getElementById('xs').onclick = () => run('share');
    }
    if (admin) $app.querySelectorAll('.card[data-id]').forEach(c => c.onclick = () => editProduct(P.find(p => p.id === c.dataset.id), () => { IB = invBarcodes(); draw(); }));
  };
  draw();
}

function stockUpdate(done) {
  const m = modal(`<h3>Update stock from file</h3><div class="sub">Pick the stock file I prepared from your PDF (stock.json). You will see a summary before anything changes.</div>
    <input id="f" type="file" accept=".json"><div id="rv"></div><button class="btn l" id="cl">Close</button>`);
  m.querySelector('#cl').onclick = () => m.remove();
  m.querySelector('#f').onchange = async e => {
    const rv = m.querySelector('#rv'); rv.innerHTML = '<div class="sub">Reading…</div>';
    try {
      const F = JSON.parse(await e.target.files[0].text()); if (!Array.isArray(F) || !F.length || !('c' in F[0]) || !('q' in F[0])) throw new Error('bad');
      const cur = await loadProducts(true), by = {}; cur.forEach(p => { by[p.id.toUpperCase()] = p; by[String(p.c).toUpperCase()] = p; });
      const seen = new Set(), chg = [], nw = [], same = [];
      F.forEach(r => { const k = String(r.c).trim().toUpperCase(), p = by[k]; if (p) { seen.add(p.id); const q = Math.round(+r.q || 0); (q !== p.q ? chg : same).push({ p, q, price: r.p }); } else nw.push(r); });
      const miss = cur.filter(p => !seen.has(p.id) && p.q !== 0), up = chg.filter(x => x.price != null && Math.abs(+x.price - x.p.p) > 0.0005);
      rv.innerHTML = `<div class="card"><div class="row"><span>Quantities that change</span><b>${chg.length}</b></div><div class="row"><span>Unchanged</span><b>${same.length}</b></div>
        <div class="row"><span>New items (not in the app)</span><b>${nw.length}</b></div><div class="row"><span>In the app but not in the file</span><b>${miss.length}</b></div>
        <div class="row"><span>Prices that differ</span><b>${F.filter(r => { const p = by[String(r.c).trim().toUpperCase()]; return p && r.p != null && Math.abs(+r.p - p.p) > 0.0005; }).length}</b></div></div>
        <label><input type="checkbox" id="o1" checked style="width:auto"> Create the ${nw.length} new items</label>
        <label><input type="checkbox" id="o2" checked style="width:auto"> Also update prices</label>
        <label><input type="checkbox" id="o3" style="width:auto"> Set the ${miss.length} items missing from the file to 0</label>
        ${chg.length ? `<div class="sub" style="margin-top:8px">Biggest changes: ${chg.slice().sort((x, y) => Math.abs(y.q - y.p.q) - Math.abs(x.q - x.p.q)).slice(0, 5).map(x => `${esc(x.p.c)} ${x.p.q}→${x.q}`).join(' • ')}</div>` : ''}
        <button class="btn g" id="ap">Apply</button>`;
      rv.querySelector('#ap').onclick = async () => {
        const b = rv.querySelector('#ap'); b.disabled = true; b.textContent = 'Updating… please wait';
        const o1 = rv.querySelector('#o1').checked, o2 = rv.querySelector('#o2').checked, o3 = rv.querySelector('#o3').checked, ops = [];
        chg.forEach(x => ops.push(['u', x.p.id, { q: x.q }])); same.forEach(x => { if (o2 && x.price != null && Math.abs(+x.price - x.p.p) > 0.0005) ops.push(['u', x.p.id, { p: +x.price }]); });
        if (o2) chg.forEach(x => { if (x.price != null && Math.abs(+x.price - x.p.p) > 0.0005) ops.find(o => o[1] === x.p.id)[2].p = +x.price; });
        if (o1) nw.forEach(r => ops.push(['s', String(r.c).trim(), { c: String(r.c).trim(), ar: r.ar || '', en: r.en || '', pk: r.pk || 1, q: Math.round(+r.q || 0), p: +r.p || 0 }]));
        if (o3) miss.forEach(p => ops.push(['u', p.id, { q: 0 }]));
        try {
          for (let i = 0; i < ops.length; i += 400) { const bt = writeBatch(db); ops.slice(i, i + 400).forEach(([t, id, d]) => t === 'u' ? bt.update(doc(db, 'products', id), d) : bt.set(doc(db, 'products', id), d)); await bt.commit(); }
          S.products = null; await loadProducts(true); m.remove(); toast(`Stock updated (${ops.length} changes)`); done();
        } catch (er) { console.error(er); toast('Update failed — please try again'); b.disabled = false; b.textContent = 'Apply'; }
      };
    } catch (er) { rv.innerHTML = '<div class="card">This is not a stock file. Use the stock.json I prepared.</div>'; }
  };
}
function modal(html) {
  const m = document.createElement('div'); m.className = 'modal'; m.innerHTML = `<div>${html}</div>`;
  m.onclick = e => { if (e.target === m) m.remove(); };
  document.body.appendChild(m); return m;
}
function editProduct(p, after) {
  const m = modal(`<h3>${esc(p.c)}</h3><div class="sub">${esc(p.en)}</div>
    <label>Stock (pieces)</label><input id="eq" type="number" inputmode="numeric" value="${p.q}">
    <label>Price (KWD)</label><input id="ep" type="number" step="0.001" inputmode="decimal" value="${p.p}">
    <label>Packing (pieces per carton)</label><input id="ek" type="number" inputmode="numeric" value="${p.pk}">
    <label>English name</label><input id="en" value="${esc(p.en)}">
    <label>Arabic name</label><input id="ea" class="ar" value="${esc(p.ar)}">
    <button class="btn l" id="hd">${isHidden(p) ? '👁 Show this item again' : '🙈 Hide this item'}</button>
    <label>Category • الفئة</label><select id="ec">${CATLIST.map(([k, n]) => `<option value="${k}" ${catOf(p) === k ? 'selected' : ''}>${n}</option>`).join('')}</select>
    ${SUBS[catOf(p)] ? `<label>Sub-category</label><select id="es">${SUBS[catOf(p)].map(([k, n]) => `<option value="${k}" ${subOf(p) === k ? 'selected' : ''}>${n}</option>`).join('')}</select>` : ''}
    <label>Barcodes</label><div id="bl"></div>
    <div class="row"><input id="nb" placeholder="Add barcode number"><button class="btn l" id="ab">Add</button></div>
    <button class="btn g" id="sv">Save</button><button class="btn l" id="cl">Close</button>`);
  const bkey = c => String(c).trim().replace(/\//g, '_');
  const drawB = () => { const L = invBarcodes()[p.c] || []; m.querySelector('#bl').innerHTML = L.length ? L.map(b => `<div class="row"><span>${esc(b.code)}</span><button class="btn l" data-x="${esc(b.id)}">✕</button></div>`).join('') : '<div class="sub">None yet</div>';
    m.querySelectorAll('[data-x]').forEach(x => x.onclick = async () => { try { await deleteDoc(doc(db, 'barcodes', x.dataset.x)); delete S.barcodes[x.dataset.x]; drawB(); } catch (e) { toast('Could not remove'); } }); };
  drawB();
  m.querySelector('#ab').onclick = async () => { const code = m.querySelector('#nb').value.trim(); if (!code) return; const old = S.barcodes[bkey(code)]; if (old && old !== p.c && !confirm('This barcode already belongs to item ' + old + '. Move it here?')) return;
    try { await setDoc(doc(db, 'barcodes', bkey(code)), { c: p.c, code, by: S.user.name, at: Date.now() }); S.barcodes[bkey(code)] = p.c; S.barcodeCodes[bkey(code)] = code; m.querySelector('#nb').value = ''; drawB(); } catch (e) { toast('Could not add'); } };
  m.querySelector('#cl').onclick = () => m.remove();
  m.querySelector('#hd').onclick = async () => { const h = !isHidden(p); try { await updateDoc(doc(db, 'products', p.id), { hidden: h }); p.hidden = h; m.remove(); toast(h ? 'Hidden' : 'Visible again'); after(); } catch (e) { toast('Could not save'); } };
  m.querySelector('#sv').onclick = async () => {
    const v = { q: Math.floor(+m.querySelector('#eq').value || 0), p: +m.querySelector('#ep').value || 0, pk: Math.max(1, +m.querySelector('#ek').value || 1), en: m.querySelector('#en').value, ar: m.querySelector('#ea').value, ...(m.querySelector('#en').value !== p.en ? { enEdited: true } : {}), ...(m.querySelector('#ea').value !== p.ar ? { arEdited: true } : {}), cat: m.querySelector('#ec').value }; if (m.querySelector('#es')) v.sub = m.querySelector('#es').value; else if (v.cat !== catOf(p)) v.sub = 'other';
    try { await updateDoc(doc(db, 'products', p.id), v); Object.assign(p, v); m.remove(); toast('Saved'); after(); } catch (e) { toast('Could not save'); }
  };
}

function editCustomer(c, after) {
  const m = modal(`<h3>${esc(c.id)}</h3>
    <label>English name</label><input id="en" value="${esc(c.en)}">
    <label>Arabic name</label><input id="ar" class="ar" value="${esc(c.ar)}">
    <label>Discount %</label><input id="ds" type="number" step="0.1" inputmode="decimal" value="${c.disc || 0}">
    <label>Status</label><select id="st"><option value="1" ${c.active !== false ? 'selected' : ''}>Active</option><option value="0" ${c.active === false ? 'selected' : ''}>Inactive</option></select>
    <button class="btn g" id="sv">Save</button><button class="btn l" id="cl">Close</button>`);
  m.querySelector('#cl').onclick = () => m.remove();
  m.querySelector('#sv').onclick = async () => {
    const v = { en: m.querySelector('#en').value, ar: m.querySelector('#ar').value, disc: +m.querySelector('#ds').value || 0, active: m.querySelector('#st').value === '1' };
    try { await updateDoc(doc(db, 'customers', c.id), v); Object.assign(c, v); m.remove(); toast('Saved'); after(); } catch (e) { toast('Could not save'); }
  };
}

async function staff() {
  $app.innerHTML = top('Staff • الموظفون') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  document.getElementById('bk').onclick = () => go('home');
  const users = (await getDocs(collection(db, 'users'))).docs.map(d => ({ id: d.id, ...d.data() }));
  const hasData = (await getDocs(collection(db, 'products'))).size > 0;
  $app.innerHTML = top('Staff • الموظفون') + `<div class="wrap">
    ${users.map(u => `<div class="card"><div class="row"><b>${esc(u.name)}</b><span class="pill ${u.active ? 'g' : 'm'}">${esc(u.role)}</span></div>
      <div class="sub">username: ${esc(u.username)} ${u.active ? '' : '• disabled'}</div>
      ${u.id !== S.user.uid ? `<button class="btn l" data-t="${u.id}" data-a="${u.active ? 0 : 1}">${u.active ? 'Disable' : 'Enable'}</button>` : ''}</div>`).join('')}
    <button class="btn" id="add">+ Add staff member</button>
    <button class="btn o" id="pin">Change my PIN</button>
    ${hasData ? '' : '<button class="btn g" id="imp">Load products &amp; customers (first time only)</button>'}</div>`;
  document.getElementById('bk').onclick = () => go('home');
  $app.querySelectorAll('[data-t]').forEach(b => b.onclick = async () => { await updateDoc(doc(db, 'users', b.dataset.t), { active: b.dataset.a === '1' }); staff(); });
  document.getElementById('pin').onclick = () => {
    const m = modal('<h3>New PIN</h3><input id="np" type="password" inputmode="numeric" maxlength="4" placeholder="4 digits"><button class="btn g" id="sv">Save</button>');
    m.querySelector('#sv').onclick = async () => { const v = m.querySelector('#np').value; if (!/^\d{4}$/.test(v)) return toast('4 digits please'); try { await updatePassword(auth.currentUser, pw(v)); m.remove(); toast('PIN changed'); } catch (e) { toast('Please log out and in again, then retry'); } };
  };
  document.getElementById('add').onclick = () => {
    const m = modal(`<h3>New staff member</h3><label>Full name</label><input id="n"><label>Username (letters/numbers, no spaces)</label><input id="u" autocapitalize="none">
      <label>PIN (4 digits)</label><input id="p" inputmode="numeric" maxlength="4"><label>Role</label>
      <select id="r"><option value="sales">Sales</option><option value="warehouse">Delivery / Warehouse</option><option value="admin">Admin</option></select>
      <button class="btn g" id="sv">Create</button><button class="btn l" id="cl">Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelector('#sv').onclick = async () => {
      const n = m.querySelector('#n').value.trim(), u = m.querySelector('#u').value.trim().toLowerCase().replace(/[^a-z0-9]/g, ''), p = m.querySelector('#p').value, r = m.querySelector('#r').value;
      if (!n || !u || !/^\d{4}$/.test(p)) return toast('Fill name, username and a 4-digit PIN');
      try {
        const sec = initializeApp(firebaseConfig, 'sec' + Date.now()); const sa = getAuth(sec);
        const c = await createUserWithEmailAndPassword(sa, mail(u), pw(p));
        await setDoc(doc(db, 'users', c.user.uid), { username: u, name: n, role: r, active: true, created: Date.now() });
        await signOut(sa); m.remove(); toast('Created'); staff();
      } catch (e) { toast(e.code === 'auth/email-already-in-use' ? 'Username already used' : 'Could not create'); }
    };
  };
  const imp = document.getElementById('imp');
  if (imp) imp.onclick = () => {
    const m = modal(`<h3>Load products &amp; customers</h3><div class="sub">Pick the two files I sent you: products.json and customers.json (select both at once).</div>
      <input id="f" type="file" accept=".json" multiple><button class="btn g" id="go">Load</button><button class="btn l" id="cl">Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelector('#go').onclick = async () => {
      const fs = [...m.querySelector('#f').files]; if (fs.length !== 2) return toast('Select both files');
      const b = m.querySelector('#go'); b.disabled = true; b.textContent = 'Loading… please wait';
      try {
        let P, C;
        for (const f of fs) { const j = JSON.parse(await f.text()); if (j[0] && 'pk' in j[0]) P = j; else C = j; }
        if (!P || !C) throw new Error('bad files');
        const all = [...P.map(p => ['products', p.c, { c: p.c, ar: p.ar, en: p.en, pk: p.pk, q: p.q, p: p.p }]),
          ...C.map(c => ['customers', c.c, { ar: c.ar, en: c.en, bal: c.bal, disc: 0, active: c.bal !== 0 }])];
        for (let i = 0; i < all.length; i += 400) {
          const bt = writeBatch(db); all.slice(i, i + 400).forEach(([col, id, d]) => bt.set(doc(db, col, id), d)); await bt.commit();
        }
        S.products = S.customers = null; m.remove(); toast(`Loaded ${P.length} products, ${C.length} customers`); staff();
      } catch (e) { console.error(e); toast('Failed — try again'); b.disabled = false; b.textContent = 'Load'; }
    };
  };
}
