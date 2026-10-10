import { getDocs, collection, query, where, orderBy } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

let X;
export function initReport(ctx) { X = ctx; }
const $ = s => document.querySelector(s);
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const parse = s => new Date(s + 'T12:00:00');
const LABEL = {
  invoice: ['🧾', 'Invoice created'], invoice_cancel: ['🚫', 'Invoice cancelled'], delivered: ['✅', 'Delivered'], delayed: ['🕒', 'Delivery delayed'],
  cust_cancelled: ['❌', 'Customer cancelled order'], reset: ['↩️', 'Delivery reset'], statement_view: ['👤', 'Opened statement'], checkin: ['📍', 'Check-in']
};

function range(period, anchor) {
  const d = parse(anchor);
  if (period === 'day') return [anchor, anchor];
  if (period === 'week') { const s = new Date(d); s.setDate(d.getDate() - d.getDay()); const e = new Date(s); e.setDate(s.getDate() + 6); return [iso(s), iso(e)]; }
  return [iso(new Date(d.getFullYear(), d.getMonth(), 1)), iso(new Date(d.getFullYear(), d.getMonth() + 1, 0))];
}
function shift(period, anchor, dir) {
  const d = parse(anchor);
  if (period === 'day') d.setDate(d.getDate() + dir); else if (period === 'week') d.setDate(d.getDate() + 7 * dir); else d.setMonth(d.getMonth() + dir, 1);
  return iso(d);
}

export async function reportHome() {
  const { S, top, esc, go, kwd } = X;
  if (S.user.role !== 'admin') return go('home');
  let period = 'day', anchor = iso(new Date()), emp = '', cache = {};
  const load = async () => {
    const [a, b] = range(period, anchor), key = a + b;
    if (!cache[key]) cache[key] = (await getDocs(query(collection(X.db, 'events'), where('date', '>=', a), where('date', '<=', b), orderBy('date', 'desc')))).docs.map(d => d.data()).sort((x, y) => y.at - x.at);
    return cache[key];
  };
  const draw = async () => {
    $('#app').innerHTML = top('Daily Report • التقرير اليومي') + '<div class="wrap"><div class="sub">Loading…</div></div>';
    $('#bk').onclick = () => go('home');
    let E; try { E = await load(); } catch (e) { console.error(e); return ($('#app').innerHTML = top('Daily Report') + '<div class="wrap"><div class="card">Could not load. If this keeps happening, tell Claude.</div></div>'); }
    const [a, b] = range(period, anchor);
    const people = [...new Map(E.map(e => [e.uid, { uid: e.uid, name: e.name, role: e.role }])).values()].sort((x, y) => x.name.localeCompare(y.name));
    const F = emp ? E.filter(e => e.uid === emp) : E;
    const cnt = (list, t) => list.filter(e => e.type === t).length;
    const inv = F.filter(e => e.type === 'invoice'), sumInv = inv.reduce((s, e) => s + (e.total || 0), 0);
    const per = people.map(p => { const L = E.filter(e => e.uid === p.uid); return { ...p, inv: cnt(L, 'invoice'), del: cnt(L, 'delivered'), dly: cnt(L, 'delayed'), can: cnt(L, 'cust_cancelled') }; });
    const hm = at => new Date(at).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    const dayStr = at => new Date(at).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
    const title = period === 'day' ? a : `${a} → ${b}`;
    $('#app').innerHTML = top('Daily Report • التقرير اليومي') + `<div class="wrap">
      <div class="tabs">${['day', 'week', 'month'].map(p => `<button data-p="${p}" class="${period === p ? 'on' : ''}">${p[0].toUpperCase() + p.slice(1)}</button>`).join('')}</div>
      <div class="row" style="margin-bottom:8px"><button class="btn l" id="pv" style="width:60px;margin:0">‹</button><b>${title}</b><button class="btn l" id="nx" style="width:60px;margin:0">›</button></div>
      <input type="date" id="dt" value="${anchor}">
      <select id="em"><option value="">All employees</option>${people.map(p => `<option value="${p.uid}" ${p.uid === emp ? 'selected' : ''}>${esc(p.name)} (${p.role})</option>`).join('')}</select>
      <div class="grid" style="margin:10px 0">
        <div class="card" style="text-align:center"><div class="sub">Invoices created</div><b style="font-size:26px">${inv.length}</b><div class="sub">${kwd(sumInv)} KWD</div></div>
        <div class="card" style="text-align:center"><div class="sub">Delivered</div><b style="font-size:26px;color:var(--g)">${cnt(F, 'delivered')}</b><div class="sub">delayed ${cnt(F, 'delayed')} • cancelled ${cnt(F, 'cust_cancelled')}</div></div></div>
      ${emp ? '' : `<h3>By employee</h3>${per.map(p => `<div class="card"><div class="row"><b>${esc(p.name)}</b><span class="pill m">${p.role}</span></div>
        <div class="sub">🧾 ${p.inv} invoices • ✅ ${p.del} delivered • 🕒 ${p.dly} delayed • ❌ ${p.can} cancelled</div></div>`).join('') || '<div class="card sub">No activity in this period.</div>'}`}
      <h3>Activity</h3>
      ${F.map(e => { const l = LABEL[e.type] || ['•', e.type], place = e.custEn || '';
        const where = e.loc ? `<a href="https://www.google.com/maps?q=${e.loc.lat},${e.loc.lng}" target="_blank" rel="noopener" style="color:var(--b);font-weight:700">📍 ${esc(place || 'Open map')}</a> <span class="sub">±${e.loc.acc}m</span>`
          : `${place ? '<b>' + esc(place) + '</b>' : ''} ${e.role === 'admin' ? '' : '<span class="sub">(no location)</span>'}`;
        return `<div class="card"><div class="row"><span>${l[0]} ${l[1]}${e.ref ? ' • ' + esc(e.ref) : ''}</span><span class="sub">${period === 'day' ? '' : dayStr(e.at) + ' '}${hm(e.at)}</span></div>
        <div>${where}</div><div class="sub">${esc(e.name)}${e.total ? ' • ' + kwd(e.total) + ' KWD' : ''}</div></div>`; }).join('') || '<div class="card sub">Nothing yet.</div>'}</div>`;
    $('#bk').onclick = () => go('home');
    document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { period = b.dataset.p; draw(); });
    $('#pv').onclick = () => { anchor = shift(period, anchor, -1); draw(); };
    $('#nx').onclick = () => { anchor = shift(period, anchor, 1); draw(); };
    $('#dt').onchange = e => { if (e.target.value) { anchor = e.target.value; draw(); } };
    $('#em').onchange = e => { emp = e.target.value; draw(); };
  };
  draw();
}
