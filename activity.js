import { collection, addDoc } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
let X;
export function initActivity(ctx) { X = ctx; }
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

export function getLoc() {
  return new Promise(res => {
    if (!navigator.geolocation) return res(null);
    navigator.geolocation.getCurrentPosition(
      p => res({ lat: +p.coords.latitude.toFixed(6), lng: +p.coords.longitude.toFixed(6), acc: Math.round(p.coords.accuracy) }),
      () => res(null), { enableHighAccuracy: true, timeout: 7000, maximumAge: 30000 });
  });
}
// Every action goes in the activity log. Location is recorded for employees only, never for admins.
export async function logEvent(type, data = {}) {
  try {
    const u = X.S.user, now = new Date();
    const loc = u.role === 'admin' ? null : await getLoc();
    await addDoc(collection(X.db, 'events'), { at: now.getTime(), date: iso(now), uid: u.uid, name: u.name, role: u.role, type, ...data, loc });
    return loc;
  } catch (e) { console.error('log', e); return null; }
}

// Manual "I am here" check-in (employees)
export function checkInModal() {
  const { S, esc, toast, loadCustomers } = X;
  const m = X.modal(`<h3>📍 Check in</h3><div class="sub">Records where you are now. Choose the customer you are visiting (optional).</div>
    <input id="q" placeholder="Search customer…"><div id="r"></div><label>Note (optional)</label><input id="nt">
    <button class="btn g" id="go">📍 Check in here</button><button class="btn l" id="cl">Close</button>`);
  let cust = null, C = [];
  loadCustomers().then(c => C = c).catch(() => {});
  const g = s => m.querySelector(s);
  g('#cl').onclick = () => m.remove();
  g('#q').oninput = e => {
    const t = e.target.value.trim().toLowerCase(); if (t.length < 2) { g('#r').innerHTML = ''; return; }
    const L = C.filter(c => c.id.includes(t) || (c.en || '').toLowerCase().includes(t) || (c.ar || '').includes(t)).slice(0, 8);
    g('#r').innerHTML = L.map(c => `<div class="card" data-id="${c.id}"><b>${esc(c.en)}</b><div class="ar sub">${esc(c.ar)}</div></div>`).join('');
    m.querySelectorAll('[data-id]').forEach(el => el.onclick = () => { cust = C.find(c => c.id === el.dataset.id); g('#q').value = cust.en; g('#r').innerHTML = ''; });
  };
  g('#go').onclick = async () => {
    const b = g('#go'); b.disabled = true; b.textContent = 'Getting location…';
    const loc = await logEvent('checkin', { cust: cust ? cust.id : '', custEn: cust ? cust.en : '', note: g('#nt').value });
    m.remove(); toast(loc ? 'Checked in ✓' : 'Saved, but location was not available (allow location access)');
  };
}
