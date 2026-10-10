import { doc, getDocs, updateDoc, collection, query, orderBy } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import { logEvent, checkInModal } from './activity.js';

let X;
export function initCalendar(ctx) { X = ctx; }
const $ = s => document.querySelector(s);
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const DAYS = [['Sun', 'أحد'], ['Mon', 'اثنين'], ['Tue', 'ثلاثاء'], ['Wed', 'أربعاء'], ['Thu', 'خميس'], ['Fri', 'جمعة'], ['Sat', 'سبت']];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

export async function calendarHome() {
  const { S, top, esc, go, kwd, toast } = X;
  $('#app').innerHTML = top('Calendar • التقويم') + '<div class="wrap"><div class="sub">Loading…</div></div>';
  $('#bk').onclick = () => go('home');
  let INV;
  try { INV = (await getDocs(query(collection(X.db, 'invoices'), orderBy('created', 'asc')))).docs.map(d => ({ id: d.id, ...d.data() })).filter(i => i.status !== 'cancelled'); }
  catch (e) { console.error(e); return ($('#app').innerHTML = top('Calendar') + '<div class="wrap"><div class="card">Could not load.</div></div>'); }
  const role = S.user.role, canAct = role === 'warehouse' || role === 'admin', showMoney = role !== 'warehouse';
  const t0 = iso(new Date()); let sel = t0, ym = t0.slice(0, 7);
  const byDay = () => { const m = {}; INV.forEach(i => (m[i.delivery] ||= []).push(i)); return m; };
  const stLabel = i => i.deliveryStatus === 'delivered' ? '<span class="pill g">✓ تم التسليم</span>' : i.deliveryStatus === 'cancelled' ? '<span class="pill r">✗ ألغاه العميل</span>' : '<span class="pill o">قيد التوصيل</span>';
  const card = i => `<div class="card" data-id="${i.id}"><div class="row"><b class="ar">${esc(i.custAr || i.custEn)}</b>${stLabel(i)}</div>
    ${i.custAr ? `<div class="sub">${esc(i.custEn)}</div>` : ''}<div class="sub">${esc(i.no)} • ${i.pieces} pcs • ${i.cartons.toFixed(1)} ctn${showMoney ? ' • ' + kwd(i.total) : ''}</div>
    ${i.note ? `<div class="sub">📝 ${esc(i.note)}</div>` : ''}${i.delays ? `<div class="sub">🕒 تم التأجيل ${i.delays.length}×</div>` : ''}</div>`;
  const draw = () => {
    const B = byDay(), [y, mo] = ym.split('-').map(Number), first = new Date(y, mo - 1, 1), n = new Date(y, mo, 0).getDate();
    const cells = []; for (let k = 0; k < first.getDay(); k++) cells.push('<div></div>');
    for (let d = 1; d <= n; d++) {
      const key = `${ym}-${String(d).padStart(2, '0')}`, L = B[key] || [], pend = L.filter(i => (i.deliveryStatus || 'pending') === 'pending').length;
      cells.push(`<button data-d="${key}" style="border:0;border-radius:10px;padding:8px 0;background:${key === sel ? 'var(--b)' : key === t0 ? '#dbe7fb' : '#fff'};color:${key === sel ? '#fff' : 'inherit'};font-size:15px">${d}${L.length ? `<br><span style="font-size:11px;font-weight:700;color:${key === sel ? '#fff' : pend ? 'var(--o)' : 'var(--g)'}">${pend || '✓'}</span>` : '<br><span style="font-size:11px">&nbsp;</span>'}</button>`);
    }
    const late = INV.filter(i => (i.deliveryStatus || 'pending') === 'pending' && i.delivery < t0).sort((a, b) => a.delivery.localeCompare(b.delivery));
    const day = B[sel] || [], cx = INV.filter(i => i.deliveryStatus === 'cancelled' && role === 'admin');
    $('#app').innerHTML = top('Calendar • التقويم') + `<div class="wrap">
      ${role === 'warehouse' ? '<button class="btn l" id="ci">📍 Check in • تسجيل الحضور</button>' : ''}
      ${role === 'admin' && cx.length ? `<div class="card" style="background:#fde7e5;color:var(--r)"><b>⚠ ${cx.length} order(s) cancelled by the customer</b><div>Open the invoice in New Invoice → Recent invoices and cancel it so the stock returns:</div>${cx.map(i => `<div>${esc(i.no)} • ${esc(i.custEn)}</div>`).join('')}</div>` : ''}
      <div class="row" style="margin-bottom:8px"><button class="btn l" id="pm" style="width:60px;margin:0">‹</button><b>${MONTHS[mo - 1]} ${y}</b><button class="btn l" id="nm" style="width:60px;margin:0">›</button></div>
      <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:4px;text-align:center;margin-bottom:12px">
        ${DAYS.map(d => `<div class="sub" style="font-size:12px">${d[0]}<br>${d[1]}</div>`).join('')}${cells.join('')}</div>
      ${late.length ? `<h3 style="color:var(--r)">متأخر • Late (${late.length})</h3>${late.map(card).join('')}` : ''}
      <h3>${sel === t0 ? 'اليوم • Today' : sel} <span class="sub">(${day.length})</span></h3>
      ${day.map(card).join('') || '<div class="card sub">لا توجد توصيلات • No deliveries</div>'}</div>`;
    $('#bk').onclick = () => go('home');
    $('#pm').onclick = () => { const d = new Date(y, mo - 2, 1); ym = iso(d).slice(0, 7); draw(); };
    $('#nm').onclick = () => { const d = new Date(y, mo, 1); ym = iso(d).slice(0, 7); draw(); };
    document.querySelectorAll('[data-d]').forEach(b => b.onclick = () => { sel = b.dataset.d; draw(); });
    document.querySelectorAll('.card[data-id]').forEach(c => c.onclick = () => open(INV.find(i => i.id === c.dataset.id)));
    const ci = $('#ci'); if (ci) ci.onclick = checkInModal;
  };
  function open(i) {
    const st = i.deliveryStatus || 'pending', admin = role === 'admin';
    const m = X.modal(`<h3 class="ar">${esc(i.custAr || i.custEn)}</h3><div class="sub">${esc(i.custEn)} • ${esc(i.no)} • ${esc(i.delivery)}</div>${stLabel(i)}
      <div style="margin:8px 0">${i.items.map(x => `<div class="row"><span>${esc(x.c)}</span><b>${x.q} <span class="sub">(${(x.q / x.pk).toFixed(x.q % x.pk ? 1 : 0)} ctn)</span></b></div><div class="sub">${esc(x.en)}</div>`).join('')}</div>
      <div class="row"><span>القطع • Pieces</span><b>${i.pieces}</b></div><div class="row"><span>الكراتين • Cartons</span><b>${i.cartons.toFixed(2)}</b></div>
      ${showMoney ? `<div class="row"><span>Total</span><b>${kwd(i.total)} KWD</b></div>` : ''}${i.note ? `<div class="sub">📝 ${esc(i.note)}</div>` : ''}
      ${canAct && st === 'pending' ? `<button class="btn g" data-a="delivered">✅ تم التسليم</button><button class="btn o" data-a="delay">🕒 تأجيل التوصيل</button><button class="btn r" data-a="cancel">❌ إلغاء من العميل</button>` : ''}
      ${admin && st !== 'pending' ? '<button class="btn o" data-a="reset">Back to pending</button>' : ''}<button class="btn l" id="cl">إغلاق • Close</button>`);
    m.querySelector('#cl').onclick = () => m.remove();
    const save = async (patch, ev, label) => {
      m.querySelectorAll('button').forEach(b => b.disabled = true);
      try {
        await logEvent(ev, { ref: i.no, cust: i.cust, custEn: i.custEn });
        await updateDoc(doc(X.db, 'invoices', i.id), patch); Object.assign(i, patch); m.remove(); toast(label); draw();
      } catch (e) { console.error(e); m.querySelectorAll('button').forEach(b => b.disabled = false); toast('Could not save'); }
    };
    m.querySelectorAll('[data-a]').forEach(b => b.onclick = () => {
      const a = b.dataset.a;
      if (a === 'delivered') { if (confirm('تأكيد التسليم؟\nConfirm delivery?')) save({ deliveryStatus: 'delivered', deliveredBy: S.user.name, deliveredAt: Date.now() }, 'delivered', 'تم التسليم ✓'); }
      else if (a === 'cancel') { if (confirm('العميل ألغى الطلب؟\nCustomer cancelled this order?')) save({ deliveryStatus: 'cancelled', deliveredBy: S.user.name, deliveredAt: Date.now() }, 'cust_cancelled', 'تم الإلغاء'); }
      else if (a === 'reset') save({ deliveryStatus: 'pending', deliveredBy: '', deliveredAt: 0 }, 'reset', 'Reset');
      else if (a === 'delay') {
        const d = X.modal(`<h3>🕒 تاريخ التوصيل الجديد</h3><input id="nd" type="date" value="${i.delivery}" min="${t0}"><button class="btn g" id="ok">حفظ • Save</button><button class="btn l" id="no">إلغاء</button>`);
        d.querySelector('#no').onclick = () => d.remove();
        d.querySelector('#ok').onclick = () => {
          const nd = d.querySelector('#nd').value; if (!nd || nd === i.delivery) return toast('اختر تاريخًا مختلفًا');
          d.remove(); sel = nd; ym = nd.slice(0, 7);
          save({ delivery: nd, delays: [...(i.delays || []), { from: i.delivery, to: nd, by: S.user.name, at: Date.now() }] }, 'delayed', 'تم التأجيل ✓');
        };
      }
    });
  }
  draw();
}
