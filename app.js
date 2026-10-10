import { firebaseConfig } from './firebase-config.js';
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import { getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, updatePassword } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { getFirestore, doc, getDoc, setDoc, getDocs, collection, writeBatch, updateDoc } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

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

import { initInvoices, invoiceHome } from './invoices.js';
import { initStatement, statementHome, chequesHome } from './statement.js';
import { initActivity, getLoc } from './activity.js';
import { initCalendar, calendarHome } from './calendar.js';
const fb = initializeApp(firebaseConfig);
const auth = getAuth(fb), db = getFirestore(fb);
const mail = u => u.trim().toLowerCase().replace(/[^a-z0-9]/g, '') + '@kitchenware.app';
const pw = pin => pin + '-kw';

const S = { user: null, products: null, customers: null, lowLimit: 5, view: 'home' };
const LOW = () => S.lowLimit;
const CTX = { db, S, top: (...a) => top(...a), esc, go: (...a) => go(...a), kwd, toast, modal: h => modal(h), loadProducts: f => loadProducts(f), loadCustomers: f => loadCustomers(f), editCustomer: (c, f) => editCustomer(c, f) };
initInvoices(CTX); initStatement(CTX); initActivity(CTX); initCalendar(CTX);

const ROLE_HOME = {
  admin: ['invoice', 'return', 'stock', 'statement', 'report', 'calendar'],
  sales: ['invoice', 'return', 'stock', 'statement', 'calendar'],
  warehouse: ['calendar']
};
const BTN = {
  invoice: ['🧾', 'New Invoice', 'فاتورة جديدة'], return: ['↩️', 'Return', 'مرتجع'], stock: ['📦', 'Stock', 'المخزون'],
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
  const fn = { home, stock, statement: statementHome, cheques: chequesHome, calendar: calendarHome, staff, invoice: invoiceHome, comingsoon }[v] || comingsoon;
  fn();
  const bk = document.getElementById('bk'); if (bk) bk.onclick = () => go('home');
  const lo = document.getElementById('lo'); if (lo) lo.onclick = () => signOut(auth);
  window.scrollTo(0, 0);
}

function home() {
  const r = S.user.role, items = ROLE_HOME[r] || [];
  const adminExtra = r === 'admin' ? `<div class="grid" style="margin-top:12px">
    <button class="big small" data-go="cheques"><span class="ic">💳</span>Cheques<span class="ar">الشيكات</span></button>
    <button class="big small" data-go="count"><span class="ic">✅</span>Stock Count<span class="ar">جرد المخزون</span></button>
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
  S.products = q.docs.map(d => ({ id: d.id, ...d.data() }));
  return S.products;
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
  await loadConfig(); let P = await loadProducts();
  let filter = 'all', q = '';
  const draw = () => {
    const low = P.filter(p => p.q > 0 && p.q <= LOW()).length, zero = P.filter(p => p.q <= 0).length;
    const t = q.trim().toLowerCase();
    let L = P.filter(p => (filter === 'all' || (filter === 'low' && p.q > 0 && p.q <= LOW()) || (filter === 'zero' && p.q <= 0)) &&
      (!t || p.c.toLowerCase().includes(t) || (p.en || '').toLowerCase().includes(t) || (p.ar || '').includes(t)));
    const shown = L.slice(0, 80);
    $app.innerHTML = top('Stock • المخزون') + `<div class="wrap">
      <input id="q" placeholder="Search item number or name…" value="${esc(q)}">
      <div class="tabs"><button data-f="all" class="${filter === 'all' ? 'on' : ''}">All ${P.length}</button>
      <button data-f="low" class="${filter === 'low' ? 'on' : ''}">Low ${low}</button>
      <button data-f="zero" class="${filter === 'zero' ? 'on' : ''}">Out ${zero}</button></div>
      ${shown.map(p => `<div class="card" data-id="${esc(p.id)}"><div class="row"><b>${esc(p.c)}</b>
        <span class="pill ${p.q <= 0 ? 'r' : p.q <= LOW() ? 'o' : 'g'}">${p.q <= 0 ? 'OUT' : p.q} ${p.q > 0 ? 'pcs' : ''}</span></div>
        <div>${esc(p.en)}</div><div class="ar sub">${esc(p.ar)}</div>
        <div class="sub">Packing ${p.pk} • ${kwd(p.p)} KWD${p.pk && p.q > 0 ? ` • ${(p.q / p.pk).toFixed(1)} cartons` : ''}</div></div>`).join('')}
      ${L.length > 80 ? `<div class="sub" style="text-align:center">Showing 80 of ${L.length}. Type to narrow down.</div>` : ''}
      ${L.length === 0 ? '<div class="card">No items found.</div>' : ''}</div>`;
    document.getElementById('bk').onclick = () => go('home');
    const qi = document.getElementById('q'); qi.oninput = e => { q = e.target.value; draw(); const n = document.getElementById('q'); n.focus(); n.setSelectionRange(q.length, q.length); };
    $app.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { filter = b.dataset.f; draw(); });
    if (admin) $app.querySelectorAll('.card[data-id]').forEach(c => c.onclick = () => editProduct(P.find(p => p.id === c.dataset.id), draw));
  };
  draw();
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
    <button class="btn g" id="sv">Save</button><button class="btn l" id="cl">Close</button>`);
  m.querySelector('#cl').onclick = () => m.remove();
  m.querySelector('#sv').onclick = async () => {
    const v = { q: Math.floor(+m.querySelector('#eq').value || 0), p: +m.querySelector('#ep').value || 0, pk: Math.max(1, +m.querySelector('#ek').value || 1), en: m.querySelector('#en').value, ar: m.querySelector('#ea').value };
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
