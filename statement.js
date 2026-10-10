import { doc, getDocs, setDoc, updateDoc, collection, query, orderBy } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

let X;
export function initStatement(ctx) { X = ctx; }
const $ = s => document.querySelector(s);
const r3 = n => Math.round(n * 1000) / 1000;
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const today = () => iso(new Date());
const dayDiff = (a, b) => Math.round((new Date(b + 'T00:00:00') - new Date(a + 'T00:00:00')) / 864e5);
const BANKS = ['NBK', 'KFH', 'Gulf Bank', 'CBK', 'Burgan Bank', 'ABK', 'Boubyan', 'Warba', 'KIB'];
export const OVERDUE_DAYS = 90;

/* ---------- data ---------- */
export async function loadLedger() {
  const [inv, chq] = await Promise.all([
    getDocs(query(collection(X.db, 'invoices'), orderBy('created', 'asc'))).then(s => s.docs.map(d => ({ id: d.id, ...d.data() }))),
    getDocs(collection(X.db, 'cheques')).then(s => s.docs.map(d => ({ id: d.id, ...d.data() })))
  ]);
  return { inv, chq };
}
// Everything about one customer's account
export function account(cust, ledger) {
  const invs = ledger.inv.filter(i => i.cust === cust.id && i.status !== 'cancelled');
  const chqs = ledger.chq.filter(c => c.cust === cust.id);
  const live = chqs.filter(c => c.status !== 'bounced');
  const paidTo = t => r3(live.reduce((a, c) => a + (c.alloc || []).filter(x => x.t === t).reduce((b, x) => b + x.a, 0), 0));
  const rows = invs.map(i => { const paid = paidTo(i.no); return { ...i, paid, rem: r3(i.total - paid) }; });
  const prev = cust.bal || 0, prevPaid = paidTo('PREV'), adv = paidTo('ADV');
  const balance = r3(prev - prevPaid + rows.reduce((a, i) => a + i.total - i.paid, 0) - adv);
  const open = rows.filter(i => i.rem > 0.0005).sort((a, b) => a.date.localeCompare(b.date));
  const oldest = open[0], age = oldest ? dayDiff(oldest.date, today()) : 0;
  return { prev, prevPaid, prevRem: r3(prev - prevPaid), adv, balance, rows, chqs, open, oldest, age, overdue: age > OVERDUE_DAYS };
}

/* ---------- customer list ---------- */
export async function statementHome() {
  const { S, top, esc, go, kwd, loadCustomers } = X;
  $('#app').innerHTML = top('Customers • العملاء') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  const C = await loadCustomers(true); let L0;
  try { L0 = await loadLedger(); } catch (e) { console.error(e); L0 = { inv: [], chq: [] }; }
  const A = new Map(C.map(c => [c.id, account(c, L0)]));
  let q = '', tab = 'all';
  const draw = () => {
    const t = q.trim().toLowerCase();
    const nOver = [...A.values()].filter(a => a.overdue).length;
    const L = C.filter(c => (c.active !== false || t) && (!t || c.id.includes(t) || (c.en || '').toLowerCase().includes(t) || (c.ar || '').includes(t)) &&
      (tab === 'all' || (tab === 'over' && A.get(c.id).overdue) || (tab === 'owe' && A.get(c.id).balance > 0.0005)))
      .sort((a, b) => tab === 'all' ? (a.en || '').localeCompare(b.en || '') : A.get(b.id).balance - A.get(a.id).balance);
    $('#app').innerHTML = top('Customers • العملاء') + `<div class="wrap"><input id="q" placeholder="Search customer…" value="${esc(q)}">
      <div class="tabs"><button data-t="all" class="${tab === 'all' ? 'on' : ''}">All</button><button data-t="owe" class="${tab === 'owe' ? 'on' : ''}">Owing</button><button data-t="over" class="${tab === 'over' ? 'on' : ''}">⚠ 90+ days ${nOver}</button></div>
      ${L.map(c => { const a = A.get(c.id); return `<div class="card" data-id="${c.id}"><div class="row"><b>${esc(c.en)}</b>
        <span class="pill ${a.balance < -0.0005 ? 'g' : a.balance > 0.0005 ? 'o' : 'm'}">${kwd(Math.abs(a.balance))}${a.balance < -0.0005 ? ' advance' : ''}</span></div>
        <div class="ar sub">${esc(c.ar) || '<i>no Arabic name</i>'}</div>
        <div class="sub">${c.id}${c.active === false ? ' • inactive' : ''} ${a.overdue ? `<span class="pill r">⚠ ${a.age} days unpaid</span>` : ''}</div></div>`; }).join('')}
      ${L.length ? '' : '<div class="card sub">No customers.</div>'}</div>`;
    $('#bk').onclick = () => go('home');
    document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; draw(); });
    const qi = $('#q'); qi.oninput = e => { q = e.target.value; draw(); const n = $('#q'); n.focus(); n.setSelectionRange(q.length, q.length); };
    document.querySelectorAll('.card[data-id]').forEach(c => c.onclick = () => detail(C.find(x => x.id === c.dataset.id), L0));
  };
  draw();
}

/* ---------- one customer statement ---------- */
async function detail(c, L0) {
  const { S, top, esc, go, kwd, toast } = X;
  let L = L0;
  const draw = () => {
    const a = account(c, L), admin = S.user.role === 'admin';
    const months = {};
    a.rows.forEach(i => { const m = i.date.slice(0, 7); (months[m] ||= { inv: 0, paid: 0 }); months[m].inv += i.total; months[m].paid += i.paid; });
    const mk = Object.keys(months).sort().reverse();
    $('#app').innerHTML = top(esc(c.en)) + `<div class="wrap">
      <div class="card" style="text-align:center"><div class="ar sub">${esc(c.ar)}</div>
        <div class="sub">${a.balance < -0.0005 ? 'Advance payment' : 'Amount owed'}</div>
        <div style="font-size:30px;font-weight:800;color:${a.balance < -0.0005 ? 'var(--g)' : a.balance > 0.0005 ? 'var(--o)' : 'var(--m)'}">${kwd(Math.abs(a.balance))} KWD</div>
        ${c.disc ? `<span class="pill g">${c.disc}% discount</span>` : ''}</div>
      ${a.overdue ? `<div class="card" style="background:#fde7e5;color:var(--r)"><b>⚠ Not paid for ${a.age} days</b><div>Oldest unpaid invoice: ${esc(a.oldest.no)} (${esc(a.oldest.date)}) — ${kwd(a.oldest.rem)} left</div></div>` : ''}
      ${admin ? '<button class="btn" id="ac">+ Add cheque</button>' : ''}
      <h3>Previous balance</h3>
      <div class="card"><div class="row"><span>Previous Balance • 10-10-2026</span><b>${kwd(a.prev)}</b></div>
        ${a.prevPaid ? `<div class="sub">Paid by cheques: ${kwd(a.prevPaid)} • left <b>${kwd(a.prevRem)}</b></div>` : ''}</div>
      <h3>By month (new invoices)</h3>
      ${mk.length ? mk.map(m => `<div class="card"><div class="row"><b>${m}</b><span class="pill ${months[m].inv - months[m].paid > 0.0005 ? 'o' : 'g'}">${kwd(months[m].inv - months[m].paid)} left</span></div>
        <div class="sub">Invoiced ${kwd(months[m].inv)} • paid ${kwd(months[m].paid)}</div></div>`).join('') : '<div class="card sub">No invoices yet.</div>'}
      <h3>Invoices</h3>
      ${a.rows.slice().reverse().map(i => `<div class="card"><div class="row"><b>${esc(i.no)}</b><span class="pill ${i.rem > 0.0005 ? 'o' : 'g'}">${i.rem > 0.0005 ? kwd(i.rem) + ' left' : 'PAID'}</span></div>
        <div class="sub">${esc(i.date)} • total ${kwd(i.total)} • paid ${kwd(i.paid)}</div></div>`).join('') || '<div class="card sub">None.</div>'}
      <h3>Cheques</h3>
      ${a.chqs.slice().sort((x, y) => y.due.localeCompare(x.due)).map(k => chequeCard(k, false)).join('') || '<div class="card sub">None.</div>'}
      ${admin ? '<button class="btn l" id="ed">Edit customer</button>' : ''}</div>`;
    $('#bk').onclick = () => statementHome();
    if (admin) {
      $('#ac').onclick = () => addCheque(c, a, async () => { L = await loadLedger(); draw(); });
      $('#ed').onclick = () => X.editCustomer(c, () => draw());
    }
  };
  draw();
}

function chequeCard(k, withCust) {
  const { esc, kwd } = X, d = dayDiff(today(), k.due);
  const st = k.status === 'bounced' ? '<span class="pill r">BOUNCED</span>' : k.status === 'cleared' ? '<span class="pill g">CLEARED</span>'
    : d > 0 ? `<span class="pill m">PENDING • due in ${d}d</span>` : `<span class="pill o">DUE ${d === 0 ? 'today' : -d + 'd ago'}</span>`;
  return `<div class="card" data-k="${k.id}"><div class="row"><b>${kwd(k.amount)}</b>${st}</div>
    ${withCust ? `<div>${esc(k.custEn)}</div>` : ''}<div class="sub">Cheque ${esc(k.no)} • ${esc(k.bank)} • due ${esc(k.due)}</div></div>`;
}

/* ---------- add cheque ---------- */
function addCheque(c, a, after) {
  const { esc, kwd, toast, S } = X;
  const targets = [];
  if (a.prevRem > 0.0005) targets.push(['PREV', `Previous balance (${kwd(a.prevRem)} left)`, a.prevRem]);
  a.open.forEach(i => targets.push([i.no, `${i.no} • ${i.date} (${kwd(i.rem)} left)`, i.rem]));
  targets.push(['ADV', 'Advance payment (not for an invoice)', 1e12]);
  let rows = [{ t: targets[0][0], a: '' }];
  const m = X.modal('');
  const draw = () => {
    m.firstElementChild.innerHTML = `<h3>New cheque • ${esc(c.en)}</h3>
      <label>Cheque number</label><input id="no" inputmode="numeric">
      <label>Bank</label><input id="bk" list="banks" placeholder="Bank"><datalist id="banks">${BANKS.map(b => `<option>${b}</option>`).join('')}</datalist>
      <label>Due date • تاريخ الاستحقاق</label><input id="du" type="date" value="${today()}">
      <label>Amount (KWD)</label><input id="am" type="number" step="0.001" inputmode="decimal">
      <div class="sub" style="margin-top:8px"><b>This cheque pays:</b></div>
      ${rows.map((r, i) => `<div class="row"><select data-t="${i}" style="flex:2">${targets.map(t => `<option value="${esc(t[0])}" ${t[0] === r.t ? 'selected' : ''}>${esc(t[1])}</option>`).join('')}</select>
        <input data-a="${i}" type="number" step="0.001" inputmode="decimal" placeholder="amount" value="${r.a}" style="flex:1"></div>`).join('')}
      <button class="btn l" id="sp">+ Split to another invoice</button>
      <button class="btn g" id="sv">Save cheque</button><button class="btn l" id="cl">Close</button>`;
    const g = s => m.querySelector(s);
    g('#cl').onclick = () => m.remove();
    g('#sp').onclick = () => { snap(); rows.push({ t: targets[0][0], a: '' }); draw(); restore(); };
    let keep = {};
    function snap() { keep = { no: g('#no').value, bk: g('#bk').value, du: g('#du').value, am: g('#am').value }; m.querySelectorAll('[data-t]').forEach(e => rows[+e.dataset.t].t = e.value); m.querySelectorAll('[data-a]').forEach(e => rows[+e.dataset.a].a = e.value); }
    function restore() { g('#no').value = keep.no; g('#bk').value = keep.bk; g('#du').value = keep.du; g('#am').value = keep.am; }
    g('#am').oninput = e => { if (rows.length === 1) { const t = targets.find(x => x[0] === g('[data-t="0"]').value), v = +e.target.value || 0; g('[data-a="0"]').value = v ? r3(Math.min(v, t[2])) : ''; } };
    m.querySelector('[data-t="0"]').onchange = () => g('#am').oninput({ target: g('#am') });
    g('#sv').onclick = async () => {
      snap();
      const amount = r3(+keep.am || 0), alloc = rows.map(r => ({ t: r.t, a: r3(+r.a || 0) })).filter(r => r.a > 0);
      if (!keep.no.trim() || !keep.bk.trim() || !keep.du || amount <= 0) return toast('Fill cheque number, bank, due date and amount');
      if (Math.abs(alloc.reduce((s, r) => s + r.a, 0) - amount) > 0.0005) return toast('The split amounts must add up to the cheque amount');
      for (const r of alloc) { const t = targets.find(x => x[0] === r.t); if (r.a > t[2] + 0.0005) return toast('More than the amount left on ' + r.t); }
      if (new Set(alloc.map(r => r.t)).size !== alloc.length) return toast('Each invoice only once');
      const ref = doc(collection(X.db, 'cheques'));
      try {
        await setDoc(ref, { cust: c.id, custEn: c.en, no: keep.no.trim(), bank: keep.bk.trim(), due: keep.du, amount, alloc, status: 'pending', by: S.user.name, created: Date.now() });
        m.remove(); toast('Cheque saved'); after();
      } catch (e) { console.error(e); toast('Could not save'); }
    };
  };
  draw();
}

/* ---------- cheque tracker (admin) ---------- */
export async function chequesHome() {
  const { S, top, esc, go, kwd, toast, loadCustomers } = X;
  $('#app').innerHTML = top('Cheques • الشيكات') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  if (S.user.role !== 'admin') return go('home');
  const [chq, C] = await Promise.all([getDocs(collection(X.db, 'cheques')).then(s => s.docs.map(d => ({ id: d.id, ...d.data() }))), loadCustomers()]);
  let tab = 'pending', q = '';
  const draw = () => {
    const t = q.trim().toLowerCase();
    const L = chq.filter(k => (tab === 'all' || k.status === tab) && (!t || (k.custEn || '').toLowerCase().includes(t) || String(k.no).includes(t) || (k.bank || '').toLowerCase().includes(t)))
      .sort((a, b) => a.due.localeCompare(b.due));
    const sum = L.reduce((s, k) => s + k.amount, 0), dueNow = chq.filter(k => k.status === 'pending' && k.due <= today()).length;
    $('#app').innerHTML = top('Cheques • الشيكات') + `<div class="wrap">
      ${dueNow ? `<div class="card" style="background:#fdeedd;color:var(--o)"><b>${dueNow} cheque(s) are due now</b> — mark them cleared or bounced.</div>` : ''}
      <input id="q" placeholder="Search customer, cheque number, bank…" value="${esc(q)}">
      <div class="tabs">${['pending', 'bounced', 'cleared', 'all'].map(x => `<button data-t="${x}" class="${tab === x ? 'on' : ''}">${x[0].toUpperCase() + x.slice(1)}</button>`).join('')}</div>
      <div class="sub">${L.length} cheques • total ${kwd(sum)} KWD</div>
      ${L.map(k => chequeCard(k, true)).join('') || '<div class="card sub">None.</div>'}</div>`;
    $('#bk').onclick = () => go('home');
    document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; draw(); });
    const qi = $('#q'); qi.oninput = e => { q = e.target.value; draw(); const n = $('#q'); n.focus(); n.setSelectionRange(q.length, q.length); };
    document.querySelectorAll('.card[data-k]').forEach(el => el.onclick = () => act(chq.find(k => k.id === el.dataset.k)));
  };
  function act(k) {
    const m = X.modal(`<h3>${kwd(k.amount)} KWD</h3><div>${esc(k.custEn)}</div><div class="sub">Cheque ${esc(k.no)} • ${esc(k.bank)} • due ${esc(k.due)}</div>
      <div class="sub">Pays: ${(k.alloc || []).map(x => `${esc(x.t === 'PREV' ? 'Previous balance' : x.t === 'ADV' ? 'Advance' : x.t)} ${kwd(x.a)}`).join(', ')}</div>
      ${k.status !== 'cleared' ? '<button class="btn g" data-s="cleared">✓ Cleared (money received)</button>' : ''}
      ${k.status !== 'bounced' ? '<button class="btn r" data-s="bounced">✗ Bounced (amount goes back on the account)</button>' : ''}
      ${k.status !== 'pending' ? '<button class="btn o" data-s="pending">Back to pending</button>' : ''}<button class="btn l" id="cl">Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    m.querySelectorAll('[data-s]').forEach(b => b.onclick = async () => {
      if (b.dataset.s === 'bounced' && !confirm('Mark this cheque as BOUNCED? The amount returns to the customer\'s account.')) return;
      try { await updateDoc(doc(X.db, 'cheques', k.id), { status: b.dataset.s, statusBy: S.user.name, statusAt: Date.now() }); k.status = b.dataset.s; m.remove(); toast('Updated'); draw(); }
      catch (e) { toast('Could not update'); }
    });
  }
  draw();
}
