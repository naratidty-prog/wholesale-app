/* ระบบขายส่งอะไหล่ + ลูกหนี้ (หน้าเว็บ) */
'use strict';

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = (v) => { const n = Number(v); return isNaN(n) ? 0 : n; };
const round2 = (n) => Math.round(n * 100) / 100;
const money = (n) => num(n).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const money0 = (n) => num(n).toLocaleString('th-TH', { maximumFractionDigits: 0 });
const normCode = (c) => String(c || '').toUpperCase().replace(/[^0-9A-Z]/g, '');
const codeKey = (c) => String(c || '').trim().toUpperCase();
const pad = (n, w) => String(n).padStart(w, '0');
function todayISO() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`; }
function addDays(iso, n) { return new Date(new Date(iso + 'T00:00:00Z').getTime() + n * 864e5).toISOString().slice(0, 10); }
function thDate(iso) { if (!iso) return ''; const [y, m, d] = String(iso).slice(0, 10).split('-'); return `${d}/${m}/${num(y) + 543}`; }
function thWeek(w) { return w ? `${thDate(w)} - ${thDate(addDays(w, 5))}` : ''; }
const truthy = (v) => v === true || v === 'TRUE' || v === 'true';
const URGENT = 'ด่วน', WEEKLY = 'รายสัปดาห์';

// แปลงตัวเลขเป็นคำอ่านภาษาไทย (เหมือน BAHTTEXT ใน Excel)
function bahtText(amount) {
  const n = round2(Math.abs(num(amount)));
  const digits = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  const units = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
  function read(s) {
    let out = '';
    if (s.length > 6) { out += read(s.slice(0, -6)) + 'ล้าน'; s = s.slice(-6); }
    const len = s.length;
    for (let i = 0; i < len; i++) {
      const d = num(s[i]), pos = len - i - 1;
      if (d === 0) continue;
      if (pos === 1 && d === 1) out += 'สิบ';
      else if (pos === 1 && d === 2) out += 'ยี่สิบ';
      else if (pos === 0 && d === 1 && len > 1 && s.slice(0, -1).replace(/0/g, '') !== '') out += 'เอ็ด';
      else out += digits[d] + units[pos];
    }
    return out;
  }
  const baht = Math.floor(n), satang = Math.round((n - baht) * 100);
  let s = baht > 0 ? read(String(baht)) + 'บาท' : '';
  s += satang > 0 ? read(String(satang)) + 'สตางค์' : (baht > 0 ? '' : 'ศูนย์บาท') + 'ถ้วน';
  return (num(amount) < 0 ? 'ลบ' : '') + s;
}

/* ================================================================ IndexedDB (เก็บรายการอะไหล่ในเครื่อง) */

const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('wholesale-app', 1);
      r.onupgradeneeded = () => { const d = r.result; d.createObjectStore('kv'); d.createObjectStore('parts', { keyPath: 'k' }); };
      r.onsuccess = () => { DB.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  tx(store, mode, fn) {
    return new Promise((res, rej) => {
      const t = DB.db.transaction(store, mode), s = t.objectStore(store);
      let out; Promise.resolve(fn(s)).then((v) => { out = v; });
      t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error);
    });
  },
  req(r) { return new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  get(store, key) { return DB.tx(store, 'readonly', (s) => DB.req(s.get(key))); },
  all(store) { return DB.tx(store, 'readonly', (s) => DB.req(s.getAll())); },
  put(store, val, key) { return DB.tx(store, 'readwrite', (s) => { key === undefined ? s.put(val) : s.put(val, key); }); },
  putMany(store, vals) { return DB.tx(store, 'readwrite', (s) => { vals.forEach((v) => s.put(v)); }); },
  clear(store) { return DB.tx(store, 'readwrite', (s) => { s.clear(); }); },
  kv: { get: (k) => DB.get('kv', k), set: (k, v) => DB.put('kv', v, k) }
};

/* ================================================================ สถานะ */

const S = {
  apiUrl: '', token: '', user: null, settings: {}, groups: [], customers: [], gemini: false,
  parts: new Map(), partsNorm: new Map(), partsList: [],
  order: null, orderWeek: null, orderDebt: null,
  dl: null, cv: null, pending: [], debtors: [], docs: null
};
const LS = {
  get(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
};

class NetError extends Error {}
async function api(action, body = {}) {
  if (!S.apiUrl) throw new NetError('ยังไม่ได้ตั้งค่าลิงก์ระบบ');
  let res;
  try {
    res = await fetch(S.apiUrl, { method: 'POST', redirect: 'follow', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ action, token: S.token }, body)) });
  } catch (e) { netChip(false); throw new NetError('ติดต่อระบบไม่ได้ (ตรวจอินเทอร์เน็ต)'); }
  if (!res.ok) { netChip(false); throw new NetError('ระบบตอบกลับผิดพลาด ' + res.status); }
  const j = await res.json();
  netChip(true);
  if (!j.ok) {
    if (/เข้าสู่ระบบใหม่/.test(j.error)) logout();
    throw new Error(j.error);
  }
  return j.data;
}
function netChip(ok) { const c = $('#net-chip'); c.textContent = ok ? '● ออนไลน์' : '● ติดต่อระบบไม่ได้'; c.style.background = ok ? 'rgba(255,255,255,.18)' : '#ffeb3b'; c.style.color = ok ? '' : '#000'; }

/* ================================================================ หน้าจอทั่วไป */

function show(id) { $$('.screen').forEach((s) => s.classList.add('hidden')); $('#' + id).classList.remove('hidden'); }
function toast(msg, bad = false, ms = 3800) {
  let t = $('#toast');
  if (!t) {
    t = document.createElement('div'); t.id = 'toast'; t.className = 'no-print';
    Object.assign(t.style, { position: 'fixed', bottom: '20px', left: '50%', transform: 'translateX(-50%)', color: '#fff', padding: '10px 20px',
      borderRadius: '8px', zIndex: 99, maxWidth: '90vw', boxShadow: '0 6px 20px rgba(0,0,0,.25)', fontSize: '15px' });
    document.body.appendChild(t);
  }
  t.style.background = bad ? '#b71c1c' : '#2b2d31';
  t.textContent = msg; t.style.display = 'block';
  clearTimeout(t._h); t._h = setTimeout(() => { t.style.display = 'none'; }, ms);
}
function openModal(html, wide = false) {
  const box = $('#modal-box');
  box.className = 'modal-box' + (wide ? ' wide' : '');
  box.innerHTML = html;
  $('#modal').classList.remove('hidden');
  const f = $('#modal-box [autofocus]') || $('#modal-box input:not([type=hidden])');
  if (f) setTimeout(() => f.focus(), 30);
  return box;
}
function closeModal() { $('#modal').classList.add('hidden'); $('#modal-box').innerHTML = ''; }
async function busy(btn, fn) {
  if (btn) btn.disabled = true;
  try { return await fn(); } catch (e) { toast(e.message, true, 6000); return undefined; } finally { if (btn) btn.disabled = false; }
}
// ชื่อผู้ทำรายการ = คนที่ล็อกอิน (หลังบ้านใช้ชื่อจากการล็อกอินเสมอ)
function staffName() {
  if (!S.user) { logout(); throw new Error('กรุณาเข้าสู่ระบบ'); }
  return S.user.name;
}
function logout() {
  S.token = ''; S.user = null; LS.del('token'); LS.del('user');
  $('#unlock-code').value = ''; show('scr-unlock');
}
// ขอรหัสแอดมินบนหน้าจอ แล้วส่งให้หลังบ้านตรวจ
function askAdmin(title, detail = '', needReason = false) {
  return new Promise((resolve) => {
    const box = openModal(`<h2>${esc(title)}</h2><p class="muted">${esc(detail)}</p>
      <form id="ap-form" autocomplete="off">
        ${needReason ? '<label>เหตุผล<input name="reason" required></label>' : ''}
        <label>รหัสแอดมิน<input name="code" type="password" required autofocus></label>
        <p class="err" id="ap-err"></p>
        <div class="foot"><button type="button" id="ap-cancel">ยกเลิก</button><button class="primary" type="submit">อนุมัติ</button></div>
      </form>`);
    $('#ap-cancel', box).onclick = () => { closeModal(); resolve(null); };
    $('#ap-form', box).onsubmit = async (ev) => {
      ev.preventDefault();
      const f = ev.target, code = f.code.value;
      try { await api('checkAdmin', { approval: code }); } catch (e) { $('#ap-err').textContent = e.message; return; }
      closeModal(); resolve({ approval: code, reason: f.reason ? f.reason.value.trim() : '' });
    };
  });
}

/* ================================================================ เริ่มต้น */

async function boot() {
  bindStatic();
  try { await DB.open(); } catch (e) { toast('เปิดที่เก็บข้อมูลในเครื่องไม่ได้', true); }
  S.apiUrl = (window.APP_CONFIG && window.APP_CONFIG.API_URL) || LS.get('apiUrl');
  S.token = LS.get('token');
  try { S.user = JSON.parse(LS.get('user') || 'null'); } catch (e) { S.user = null; }
  if (!S.user) S.token = '';
  $('#login-user').value = LS.get('lastUser');
  if (!S.apiUrl) return show('scr-setup');
  if (!S.token) return show('scr-unlock');
  await start();
}
async function start() {
  $('#staff-name').textContent = S.user ? S.user.name : '';
  show('scr-app');
  try {
    const b = await api('boot');
    S.settings = b.settings; S.groups = b.groups.filter((g) => g.id); S.customers = b.customers.filter((c) => c.id); S.gemini = b.gemini;
  } catch (e) { toast(e.message, true, 8000); if (!S.token) return; }
  await loadPartsFromDb();
  syncParts();
  newOrder();
  mountPickers();
  if (!$('#dc-from').value) { $('#dc-from').value = todayISO(); $('#dc-to').value = todayISO(); }
}

/* ================================================================ อะไหล่ */

function partFromRow(r) { return { k: codeKey(r[0]), code: r[0], name: r[1], model: r[2], cost: round2(num(r[3])), price: round2(num(r[4])), u: r[5], status: r[6] || '' }; }
async function loadPartsFromDb() {
  const rows = DB.db ? await DB.all('parts') : [];
  S.parts = new Map(rows.map((p) => [p.k, p]));
  S.partsNorm = new Map();
  rows.forEach((p) => { const n = normCode(p.code); if (n && !S.partsNorm.has(n)) S.partsNorm.set(n, p); });
  S.partsList = rows;
  $('#parts-chip').textContent = `อะไหล่ ${rows.length.toLocaleString()} รายการ`;
}
async function syncParts(full = false) {
  if (!DB.db) return;
  const since = full ? '' : await DB.kv.get('partsSince');
  const bar = $('#parts-progress');
  try {
    if (!since || !S.partsList.length) {
      let offset = 0, total = 0, maxU = '';
      bar.classList.remove('hidden');
      await DB.clear('parts');
      for (;;) {
        bar.textContent = `กำลังโหลดรายการอะไหล่จากระบบขายอะไหล่ 5 สาขา ${offset.toLocaleString()} / ${total ? total.toLocaleString() : '...'} (ครั้งแรกครั้งเดียว)`;
        const d = await api('parts', { offset, limit: 10000 });
        total = d.total;
        const items = d.rows.map(partFromRow);
        items.forEach((p) => { if (p.u > maxU) maxU = p.u; });
        await DB.putMany('parts', items);
        offset += 10000;
        if (d.done) break;
      }
      await DB.kv.set('partsSince', maxU || '0');
      bar.classList.add('hidden');
    } else {
      const d = await api('parts', { since });
      if (d.rows.length) {
        const items = d.rows.map(partFromRow);
        let maxU = since; items.forEach((p) => { if (p.u > maxU) maxU = p.u; });
        await DB.putMany('parts', items); await DB.kv.set('partsSince', maxU);
      }
    }
    await loadPartsFromDb();
  } catch (e) {
    bar.classList.add('hidden');
    toast('โหลดรายการอะไหล่ไม่สำเร็จ: ' + e.message, true, 8000);
  }
}
function findPart(code) { return S.parts.get(codeKey(code)) || (normCode(code) ? S.partsNorm.get(normCode(code)) : undefined); }
// รหัสที่อ่านจากรูปผิดไปตัวเดียว เช่น O/0, I/1, S/5, B/8, Z/2 ถ้าแก้แล้วตรงรหัสเดียวพอดีให้ใช้รหัสนั้น
const LOOKALIKE = { O: '0D', 0: 'OD', D: '0O', I: '1L', 1: 'IL7', L: '1I', S: '5', 5: 'S', B: '8', 8: 'B3', Z: '2', 2: 'Z', G: '6', 6: 'G', T: '7', 7: 'T1', Q: '0', U: 'V', V: 'U', 3: '8' };
function fuzzyPart(code) {
  const n = normCode(code);
  if (n.length < 5) return null;
  const hits = new Map();
  for (let i = 0; i < n.length; i++) {
    for (const ch of (LOOKALIKE[n[i]] || '')) {
      const p = S.partsNorm.get(n.slice(0, i) + ch + n.slice(i + 1));
      if (p) hits.set(p.k, p);
    }
  }
  return hits.size === 1 ? [...hits.values()][0] : null;
}
function searchParts(q, limit = 80) {
  q = q.trim().toLowerCase();
  if (!q) return [];
  const words = q.split(/\s+/), nq = normCode(q), out = [];
  for (const p of S.partsList) {
    const hay = (p.code + ' ' + p.name + ' ' + p.model).toLowerCase();
    if ((nq.length >= 3 && normCode(p.code).includes(nq)) || words.every((w) => hay.includes(w))) { out.push(p); if (out.length >= limit) break; }
  }
  return out;
}
function openPartSearch(q, onPick) {
  const box = openModal(`<h2>ค้นหาอะไหล่</h2><input id="ps-q" value="${esc(q)}" placeholder="ชื่อ / รหัส / รุ่น" autofocus>
    <div class="table-wrap tall" style="margin-top:10px"><table class="lines"><thead><tr><th>รหัส</th><th>ชื่อ</th><th>รุ่น</th><th class="num">ราคาขาย</th></tr></thead><tbody id="ps-body"></tbody></table></div>
    <div class="foot"><button id="ps-close">ปิด</button></div>`, true);
  const render = () => {
    const rows = searchParts($('#ps-q').value);
    $('#ps-body').innerHTML = rows.map((p, i) => `<tr class="clickable" data-i="${i}"><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td>${esc(p.model)}</td><td class="num">${money(p.price)}</td></tr>`).join('')
      || '<tr><td colspan="4" class="muted">ไม่พบ</td></tr>';
    $$('#ps-body tr.clickable').forEach((tr) => { tr.onclick = () => { closeModal(); onPick(rows[tr.dataset.i]); }; });
  };
  $('#ps-q', box).oninput = render; $('#ps-close', box).onclick = closeModal; render();
}

/* ================================================================ ลูกค้า */

const custById = (id) => S.customers.find((c) => c.id === id);
const custLabel = (c) => c ? `${c.code} · ${c.display}` : '';
const groupById = (id) => S.groups.find((g) => g.id === id) || S.groups.find((g) => g.id === 'GENERAL') || { mode: 'discount', low_pct: 25, high_pct: 29, threshold: 50000 };
function custFull(c) { return c ? [c.title, c.first, c.last].filter(Boolean).join(' ').trim() : ''; }
function searchCustomers(q) {
  q = q.trim().toLowerCase();
  const list = S.customers.filter((c) => c.active !== false && c.active !== 'FALSE');
  if (!q) return list.slice(0, 60);
  return list.filter((c) => (c.code + ' ' + c.display + ' ' + custFull(c) + ' ' + c.phone).toLowerCase().includes(q)).slice(0, 60);
}
const PICK_HANDLERS = { order: (c) => pickOrderCustomer(c), deliver: (c) => pickDeliverCustomer(c), customer: (c) => loadCustomerView(c && c.id) };
const PICKED = {};
function mountPickers() { $$('.cust-pick').forEach((el) => renderPicker(el)); }
function setPicked(scope, c) { PICKED[scope] = c || null; renderPicker($(`.cust-pick[data-for="${scope}"]`)); }
function renderPicker(el) {
  const scope = el.dataset.for, c = PICKED[scope];
  if (c) {
    const g = groupById(c.group);
    el.innerHTML = `<div class="chosen"><div><b>${esc(c.display)}</b><div class="small muted">รหัส ${esc(c.code)} · ${esc(g.name || c.group)}${c.phone ? ' · ' + esc(c.phone) : ''}</div></div><button class="ghost mini">เปลี่ยน</button></div>`;
    $('button', el).onclick = () => { PICKED[scope] = null; renderPicker(el); PICK_HANDLERS[scope](null); $('input', el).focus(); };
    return;
  }
  el.innerHTML = `<input placeholder="พิมพ์ชื่อ / รหัส / เบอร์ลูกค้า" autocomplete="off"><div class="list hidden"></div>`;
  const inp = $('input', el), list = $('.list', el);
  let rows = [], sel = 0;
  const draw = () => {
    rows = searchCustomers(inp.value);
    list.innerHTML = rows.map((x, i) => `<div data-i="${i}" class="${i === sel ? 'sel' : ''}"><span class="code">${esc(x.code)}</span><span>${esc(x.display)}</span></div>`).join('') || '<div class="muted">ไม่พบลูกค้า</div>';
    list.classList.remove('hidden');
    $$('div[data-i]', list).forEach((d) => { d.onmousedown = (ev) => { ev.preventDefault(); choose(rows[d.dataset.i]); }; });
  };
  const choose = (x) => { if (!x) return; list.classList.add('hidden'); setPicked(scope, x); PICK_HANDLERS[scope](x); };
  inp.onfocus = draw;
  inp.oninput = () => { sel = 0; draw(); };
  inp.onblur = () => setTimeout(() => list.classList.add('hidden'), 150);
  inp.onkeydown = (ev) => {
    if (ev.key === 'ArrowDown') { sel = Math.min(sel + 1, rows.length - 1); draw(); ev.preventDefault(); }
    else if (ev.key === 'ArrowUp') { sel = Math.max(sel - 1, 0); draw(); ev.preventDefault(); }
    else if (ev.key === 'Enter') { ev.preventDefault(); choose(rows[sel]); }
  };
}

/* ---------------------------------------------------------------- ราคา (ฝั่งหน้าจอใช้ประมาณ หลังบ้านคิดจริง) */

function rateFor(c, type, weekGross, forced) {
  const g = groupById(c.group), s = S.settings;
  if (type === URGENT) return { pct: num(s.urgent_pct || 20), desc: `ด่วน ลด ${num(s.urgent_pct || 20)}%` };
  if (g.mode === 'costplus') return { cp: num(c.cp_pct), up: truthy(c.cp_roundup), desc: `ทุนรวม VAT${truthy(c.cp_roundup) ? ' (ปัดขึ้น)' : ''} + ${num(c.cp_pct)}%` };
  const high = weekGross >= num(g.threshold) || forced, pct = high ? num(g.high_pct) : num(g.low_pct);
  return { pct, high, desc: `รายสัปดาห์ ลด ${pct}%` };
}
function unitNet(l, r) {
  if (l.manual) return num(l.price);
  if (r.cp != null) { if (!num(l.cost)) return num(l.price) * 0.75; let b = num(l.cost) * 1.07; if (r.up) b = Math.ceil(b); return b * (1 + r.cp / 100); }
  return num(l.price) * (1 - r.pct / 100);
}
const lineNet = (l, r) => Math.round(num(l.qty) * unitNet(l, r) + 1e-9);

/* ================================================================ รับออเดอร์ */

function newOrder() {
  S.order = { type: WEEKLY, lines: [], images: [], source: '' };
  $('#order-deposit').value = ''; $('#order-dep-ref').value = ''; $('#order-note').value = ''; $('#order-dep-method').value = 'เงินสด';
  $('#order-images').innerHTML = ''; $('#order-images').classList.add('hidden'); $('#order-review').classList.add('hidden');
  $$('#order-type button').forEach((b) => b.classList.toggle('on', b.dataset.v === WEEKLY));
  renderOrder();
}
async function pickOrderCustomer(c) {
  S.order.customer = c; S.orderWeek = null; S.orderDebt = null;
  $('#order-lock').classList.add('hidden');
  renderOrder();
  if (!c) return;
  try {
    const [w, cv] = await Promise.all([api('week', { customer_id: c.id }), api('customer', { customer_id: c.id, days: 0 })]);
    if (S.order.customer !== c) return;
    S.orderWeek = w; S.orderDebt = cv.debt;
    if (cv.debt.locked) { $('#order-lock').innerHTML = `⛔ ลูกค้ารายนี้ค้างชำระเกินกำหนด ${cv.debt.late} วัน (ยอดหนี้ ${money(cv.debt.balance)}) ถูกล็อกการขาย ต้องให้แอดมินอนุมัติ`; $('#order-lock').classList.remove('hidden'); }
    else if (cv.debt.overdue > 0) { $('#order-lock').innerHTML = `⚠️ มียอดเกินกำหนดชำระ ${money(cv.debt.overdue)} บาท (เกิน ${cv.debt.late} วัน)`; $('#order-lock').classList.remove('hidden'); }
  } catch (e) { toast(e.message, true); }
  renderOrder();
}
function orderRate() {
  const o = S.order, c = o.customer;
  if (!c) return { pct: 0, desc: '' };
  const cartGross = o.lines.reduce((t, l) => t + (l.manual ? 0 : num(l.qty) * num(l.price)), 0);
  const w = S.orderWeek;
  return rateFor(c, o.type, (w ? w.gross : 0) + cartGross, w && w.forced);
}
function renderWeekBox() {
  const box = $('#week-box'), o = S.order, c = o.customer;
  if (!c) { box.innerHTML = '<p class="muted">เลือกลูกค้าเพื่อดูยอดสั่งสัปดาห์นี้</p>'; return; }
  const g = groupById(c.group), r = orderRate();
  if (o.type === URGENT) { box.innerHTML = `<h3>สั่งด่วน (ได้ของ 1-3 วัน)</h3><div class="week-rate">${esc(r.desc)}</div><p class="muted small">ไม่นับรวมยอดสั่งรายสัปดาห์</p>`; return; }
  if (g.mode === 'costplus') { box.innerHTML = `<h3>${esc(g.name)}</h3><div class="week-rate">${esc(r.desc)}</div><p class="muted small">คิดจากราคาทุนรวม VAT ไม่ขึ้นกับยอดสั่ง</p>`; return; }
  const w = S.orderWeek, before = w ? w.gross : 0;
  const cart = o.lines.reduce((t, l) => t + (l.manual ? 0 : num(l.qty) * num(l.price)), 0), total = before + cart, th = num(g.threshold) || 1;
  const left = Math.max(0, th - total);
  box.innerHTML = `<h3>ยอดสั่งรายสัปดาห์ ${w ? esc(thWeek(w.week)) : ''} <span class="muted small">(ปิดรับเสาร์ ${esc(S.settings.week_cutoff || '14:30')} น.)</span></h3>
    <div class="row wrap"><div class="grow"><div class="small muted">สั่งไปแล้ว ${money(before)} + ใบนี้ ${money(cart)} = <b>${money(total)}</b> / ${money0(th)}</div>
    <div class="progress"><div style="width:${Math.min(100, total / th * 100)}%"></div></div>
    <div class="small">${left > 0 ? `อีก <b class="late">${money(left)}</b> บาท ได้ส่วนลด ${num(g.high_pct)}%` : `✅ ถึงยอดแล้ว ได้ส่วนลด ${num(g.high_pct)}%`}${w && w.forced ? ' <span class="tag ok">แอดมินอนุมัติอัตราสูงแล้ว</span>' : ''}</div></div>
    <div class="week-rate">ลด ${r.pct}%</div></div>`;
}
function renderOrder() {
  const o = S.order, r = orderRate(), tb = $('#order-table tbody');
  tb.innerHTML = o.lines.map((l, i) => {
    const cls = l.flag === 'bad' ? 'bad' : l.flag === 'warn' ? 'warn' : '';
    const codeCell = l.flag === 'bad' ? `<input class="fix" data-i="${i}" value="${esc(l.code)}" placeholder="แก้รหัส"><div class="small">${esc(l.msg || '')}</div>`
      : `<b>${esc(l.code)}</b>${l.manual ? ' <span class="tag">พิมพ์เอง</span>' : ''}${l.msg ? `<div class="small" style="color:var(--warn)">${esc(l.msg)}</div>` : ''}`;
    const net = l.flag === 'bad' || !o.customer ? '' : money(lineNet(l, r));
    const unit = l.flag === 'bad' || !o.customer ? '' : money(unitNet(l, r));
    return `<tr class="${cls}"><td>${i + 1}</td><td>${codeCell}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td>
      <td class="num"><input class="qty" type="number" min="0" data-i="${i}" value="${l.qty}"></td><td class="num">${l.flag === 'bad' ? '' : money(l.price)}</td>
      <td class="num">${unit}</td><td class="num">${net}</td>
      <td>${l.flag === 'bad' ? `<button class="mini" data-find="${i}">ค้นหา</button>` : ''}<button class="mini danger" data-del="${i}">ลบ</button></td></tr>`;
  }).join('') || '<tr><td colspan="9" class="muted">ยังไม่มีรายการ ยิงบาร์โค้ด พิมพ์รหัส หรือกด "อ่านจากรูป / PDF"</td></tr>';
  $$('#order-table input.qty').forEach((x) => { x.onchange = () => { o.lines[x.dataset.i].qty = Math.max(0, num(x.value)); renderOrder(); }; });
  $$('#order-table input.fix').forEach((x) => { x.onchange = () => fixLine(num(x.dataset.i), x.value); });
  $$('#order-table [data-del]').forEach((b) => { b.onclick = () => { o.lines.splice(num(b.dataset.del), 1); renderOrder(); }; });
  $$('#order-table [data-find]').forEach((b) => { b.onclick = () => { const i = num(b.dataset.find); openPartSearch(o.lines[i].name || '', (p) => setLinePart(i, p, '')); }; });
  const good = o.lines.filter((l) => l.flag !== 'bad' && num(l.qty) > 0);
  const gross = good.reduce((t, l) => t + num(l.qty) * num(l.price), 0), net = o.customer ? good.reduce((t, l) => t + lineNet(l, r), 0) : 0;
  $('#ot-count').textContent = good.length + (o.lines.length > good.length ? ` (ต้องแก้ ${o.lines.length - good.length})` : '');
  $('#ot-gross').textContent = money(gross); $('#ot-disc').textContent = money(o.customer ? gross - net : 0); $('#ot-net').textContent = money(net);
  $('#ot-dep').textContent = money($('#order-deposit').value);
  const bad = o.lines.filter((l) => l.flag === 'bad').length, warn = o.lines.filter((l) => l.flag === 'warn').length;
  const rv = $('#order-review');
  if (o.source === 'รูป' && (bad || warn)) { rv.innerHTML = `ตรวจรายการที่อ่านจากรูปก่อนบันทึก: <span class="tag red">แดง ${bad} รายการ ไม่พบรหัส ต้องแก้</span> <span class="tag warn">เหลือง ${warn} รายการ ระบบเดาให้ ตรวจกับรูป</span>`; rv.classList.remove('hidden'); }
  else rv.classList.add('hidden');
  renderWeekBox();
}
function setLinePart(i, p, msg) {
  const l = S.order.lines[i];
  Object.assign(l, { code: p.code, name: p.name, model: p.model, price: p.price, cost: p.cost, manual: false, flag: msg ? 'warn' : '', msg });
  renderOrder();
}
function fixLine(i, code) {
  const p = findPart(code);
  if (p) setLinePart(i, p, '');
  else { S.order.lines[i].code = code; S.order.lines[i].msg = 'ไม่พบรหัสนี้ในรายการอะไหล่'; renderOrder(); toast('ไม่พบรหัส ' + code, true); }
}
function addOrderPart(p, qty) {
  const ex = S.order.lines.find((l) => !l.manual && l.flag !== 'bad' && l.code === p.code);
  if (ex) ex.qty = num(ex.qty) + qty;
  else S.order.lines.push({ code: p.code, name: p.name, model: p.model, price: p.price, cost: p.cost, qty, manual: false, flag: '' });
  renderOrder();
}
function orderScan(text) {
  let raw = String(text || '').trim(), qty = Math.max(1, num($('#order-qty').value) || 1);
  if (!raw) return;
  const m = raw.match(/^(\d+)\s*[*xX]\s*(.+)$/);
  if (m) { qty = num(m[1]); raw = m[2].trim(); }
  const p = findPart(raw);
  if (p) { addOrderPart(p, qty); $('#order-qty').value = 1; return; }
  const f = fuzzyPart(raw);
  if (f) { S.order.lines.push({ code: f.code, name: f.name, model: f.model, price: f.price, cost: f.cost, qty, manual: false, flag: 'warn', msg: `พิมพ์ "${raw}" เดาเป็น ${f.code}` }); renderOrder(); return; }
  openPartSearch(raw, (x) => addOrderPart(x, qty));
}
function manualLine() {
  const box = openModal(`<h2>รายการพิมพ์เอง</h2><p class="muted">สำหรับของที่ไม่มีในรายการอะไหล่ เช่น น้ำมัน ราคาที่ใส่คือราคาสุทธิ (ไม่คิดส่วนลด)</p>
    <form id="ml-form" class="form2"><label>รหัส (ถ้ามี)<input name="code"></label><label>รุ่น<input name="model"></label>
    <label class="span2">ชื่อรายการ<input name="name" required autofocus></label><label>จำนวน<input name="qty" type="number" min="1" value="1" required></label>
    <label>ราคาสุทธิต่อชิ้น<input name="price" type="number" min="0" step="0.01" required></label>
    <div class="foot span2"><button type="button" id="ml-cancel">ยกเลิก</button><button class="primary">เพิ่ม</button></div></form>`);
  $('#ml-cancel', box).onclick = closeModal;
  $('#ml-form', box).onsubmit = (ev) => {
    ev.preventDefault(); const f = ev.target;
    S.order.lines.push({ code: f.code.value.trim(), name: f.name.value.trim(), model: f.model.value.trim(), qty: num(f.qty.value), price: num(f.price.value), cost: 0, manual: true, flag: '' });
    closeModal(); renderOrder();
  };
}

/* ---------------------------------------------------------------- อ่านรูป / PDF */

function fileToBase64(file) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = () => rej(r.error); r.readAsDataURL(file); });
}
// ย่อรูปให้ไม่เกิน 2000px ส่งเร็วขึ้น แต่ยังอ่านตัวอักษรได้
function shrinkImage(file) {
  return new Promise((res) => {
    const url = URL.createObjectURL(file), img = new Image();
    img.onload = () => {
      const k = Math.min(1, 2000 / Math.max(img.width, img.height)), c = document.createElement('canvas');
      c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      res({ mime: 'image/jpeg', data: c.toDataURL('image/jpeg', 0.88).split(',')[1], url });
    };
    img.onerror = async () => res({ mime: file.type, data: await fileToBase64(file), url });
    img.src = url;
  });
}
async function readFiles(files) {
  if (!files.length) return;
  if (!S.gemini) { toast('ยังไม่ได้ใส่รหัส Gemini ในระบบหลังบ้าน (ดูคู่มือ ขั้นตอนอ่านรูป)', true, 8000); return; }
  const wrap = $('#order-images'); wrap.classList.remove('hidden');
  S.order.source = 'รูป';
  for (const file of files) {
    const fig = document.createElement('figure');
    const isPdf = file.type === 'application/pdf';
    if (isPdf && file.size > 15e6) { toast(file.name + ' ใหญ่เกิน 15MB', true); continue; }
    const pay = isPdf ? { mime: 'application/pdf', data: await fileToBase64(file), url: URL.createObjectURL(file) } : await shrinkImage(file);
    fig.innerHTML = (isPdf ? `<a class="pdf" href="${pay.url}" target="_blank">PDF</a>` : `<img src="${pay.url}">`) + `<figcaption>${esc(file.name)}<br><span class="st">กำลังอ่าน...</span></figcaption>`;
    wrap.appendChild(fig);
    const im = $('img', fig); if (im) im.onclick = () => openModal(`<div class="zoom"><img src="${pay.url}"></div><div class="foot"><button onclick="closeModal()">ปิด</button></div>`, true);
    try {
      const out = await api('readImage', { mime: pay.mime, data: pay.data });
      let ok = 0, bad = 0;
      out.lines.forEach((x) => {
        const qty = Math.max(1, num(x.qty) || 1), raw = x.code;
        let p = raw ? findPart(raw) : null, msg = '';
        if (p && x.unsure) msg = `อ่านไม่ชัด: "${raw}"`;
        if (!p && raw) { p = fuzzyPart(raw); if (p) msg = `อ่านได้ "${raw}" เดาเป็น ${p.code}`; }
        if (p) { S.order.lines.push({ code: p.code, name: p.name, model: p.model, price: p.price, cost: p.cost, qty, manual: false, flag: msg ? 'warn' : '', msg, raw }); ok++; }
        else { S.order.lines.push({ code: raw || '', name: x.name || '', model: '', price: 0, cost: 0, qty, manual: false, flag: 'bad', msg: raw ? `อ่านได้ "${raw}" ไม่พบในรายการ` : 'ไม่มีรหัส', raw }); bad++; }
      });
      $('.st', fig).textContent = `อ่านได้ ${out.lines.length} รายการ${bad ? ` (ต้องแก้ ${bad})` : ''}`;
      if (!S.order.customer && out.customer) {
        const hit = searchCustomers(out.customer);
        if (hit.length === 1) { setPicked('order', hit[0]); pickOrderCustomer(hit[0]); toast('เลือกลูกค้า ' + hit[0].display + ' ตามที่เขียนในรูป กรุณาตรวจ'); }
      }
    } catch (e) { $('.st', fig).textContent = 'อ่านไม่สำเร็จ'; toast(e.message, true, 8000); }
    renderOrder();
  }
}

async function saveOrder(btn) {
  const o = S.order;
  await busy(btn, async () => {
    const by = staffName();
    if (!o.customer) throw new Error('กรุณาเลือกลูกค้า');
    if (o.lines.some((l) => l.flag === 'bad')) throw new Error('ยังมีรายการสีแดงที่ไม่พบรหัส แก้หรือลบก่อน');
    const lines = o.lines.filter((l) => num(l.qty) > 0).map((l) => ({ code: l.code, qty: num(l.qty), manual: l.manual, name: l.name, model: l.model, price: l.price }));
    if (!lines.length) throw new Error('ยังไม่มีรายการ');
    const body = { by, order: { customer_id: o.customer.id, type: o.type, lines, deposit: num($('#order-deposit').value), deposit_method: $('#order-dep-method').value,
      deposit_ref: $('#order-dep-ref').value, note: $('#order-note').value, source: o.source } };
    if (S.orderDebt && S.orderDebt.locked) {
      const ap = await askAdmin('ลูกค้าถูกล็อกการขาย', `ค้างชำระเกินกำหนด ${S.orderDebt.late} วัน ต้องให้แอดมินอนุมัติรับออเดอร์`);
      if (!ap) return;
      body.approval = ap.approval;
    }
    const saved = await api('saveOrder', body);
    toast('บันทึกออเดอร์ ' + saved.order_no + ' แล้ว');
    printDocs(orderDoc(saved, 'สำหรับร้าน') + orderDoc(saved, 'สำหรับลูกค้า'));
    const c = o.customer;
    newOrder(); S.order.customer = c; pickOrderCustomer(c);
  });
}

/* ================================================================ ส่งมอบ */

async function pickDeliverCustomer(c) {
  S.dl = c ? { customer: c, rows: [], debt: null } : null;
  $('#deliver-lock').classList.add('hidden');
  renderDeliver();
  if (!c) return;
  await busy(null, async () => {
    const cv = await api('customer', { customer_id: c.id, days: 0 });
    if (!S.dl || S.dl.customer !== c) return;
    S.dl.rows = cv.outstanding.map((r) => Object.assign(r, { send: 0 }));
    S.dl.debt = cv.debt; S.dl.deposits = cv.deposits;
    if (cv.debt.locked) { $('#deliver-lock').innerHTML = `⛔ ค้างชำระเกินกำหนด ${cv.debt.late} วัน ถูกล็อกการขาย ต้องให้แอดมินอนุมัติ`; $('#deliver-lock').classList.remove('hidden'); }
    else if (cv.debt.balance > 0) { $('#deliver-lock').innerHTML = `ยอดหนี้คงค้าง ${money(cv.debt.balance)} บาท${cv.debt.overdue > 0 ? ` (เกินกำหนด ${money(cv.debt.overdue)})` : ''}`; $('#deliver-lock').className = 'alert info'; }
    renderDeliver();
    $('#deliver-scan').focus();
  });
}
function renderDeliver() {
  const tb = $('#deliver-table tbody'), d = S.dl;
  if (!d) { tb.innerHTML = '<tr><td colspan="9" class="muted">เลือกลูกค้าก่อน</td></tr>'; deliverTotals(); return; }
  if (!d.rows.length) { tb.innerHTML = '<tr><td colspan="9" class="muted">ลูกค้ารายนี้ไม่มีอะไหล่ค้างรับ</td></tr>'; deliverTotals(); return; }
  let last = '';
  tb.innerHTML = d.rows.map((r, i) => {
    const head = r.order_no !== last ? `<tr class="group"><td colspan="9">${esc(r.order_no)} · ${thDate(r.date)} · ${r.type === URGENT ? '<span class="tag red">สั่งด่วน</span>' : `<span class="tag">รายสัปดาห์ ${esc(thWeek(r.week))}</span>`}${r.deposit_left > 0 ? ` <span class="tag ok">มัดจำเหลือ ${money(r.deposit_left)}</span>` : ''}</td></tr>` : '';
    last = r.order_no;
    return head + `<tr class="${r.send > 0 ? 'warn' : ''}"><td class="muted small">${esc(r.order_no)}</td><td><b>${esc(r.code)}</b></td><td>${esc(r.name)}</td><td class="num">${r.qty}</td><td class="num">${r.left}</td>
      <td class="num"><input class="qty" type="number" min="0" max="${r.left}" data-i="${i}" value="${r.send || ''}"></td><td class="num">${money(r.unit_net)}</td>
      <td class="num">${r.send ? money(Math.round(r.send * r.unit_net + 1e-9)) : ''}</td><td class="small muted">${esc(r.rate)}</td></tr>`;
  }).join('');
  $$('#deliver-table input.qty').forEach((x) => { x.onchange = () => { const r = d.rows[x.dataset.i]; r.send = Math.min(r.left, Math.max(0, num(x.value))); renderDeliver(); }; });
  deliverTotals();
}
function deliverTotals() {
  const d = S.dl, rows = d ? d.rows.filter((r) => r.send > 0) : [];
  const gross = rows.reduce((t, r) => t + r.send * r.price, 0), net = rows.reduce((t, r) => t + Math.round(r.send * r.unit_net + 1e-9), 0);
  const orders = new Set(rows.map((r) => r.order_no));
  const depAv = d ? d.rows.filter((r, i, a) => a.findIndex((x) => x.order_no === r.order_no) === i && orders.has(r.order_no)).reduce((t, r) => t + num(r.deposit_left), 0) : 0;
  const dep = $('#dl-usedep').checked ? Math.min(depAv, net) : 0;
  $('#dt-gross').textContent = money(gross); $('#dt-disc').textContent = money(gross - net); $('#dt-net').textContent = money(net);
  $('#dt-depav').textContent = money(depAv); $('#dt-dep').textContent = money(dep); $('#dt-due').textContent = money(net - dep);
}
function deliverScan(text) {
  const d = S.dl, raw = String(text || '').trim();
  if (!raw) return;
  if (!d) { toast('เลือกลูกค้าก่อน', true); return; }
  const p = findPart(raw), key = p ? p.code : codeKey(raw), n = normCode(key);
  const r = d.rows.find((x) => normCode(x.code) === n && x.send < x.left);
  if (r) { r.send += 1; renderDeliver(); toast(`${r.code} ส่ง ${r.send}/${r.left}`); return; }
  const done = d.rows.find((x) => normCode(x.code) === n);
  beep();
  toast(done ? `${key} ครบจำนวนค้างรับแล้ว` : `${key} ไม่มีในออเดอร์ค้างรับของลูกค้านี้`, true);
}
function beep() { try { const a = new AudioContext(), o = a.createOscillator(); o.frequency.value = 330; o.connect(a.destination); o.start(); setTimeout(() => { o.stop(); a.close(); }, 250); } catch (e) { /* ignore */ } }
async function saveDelivery(btn) {
  const d = S.dl;
  await busy(btn, async () => {
    const by = staffName();
    if (!d) throw new Error('เลือกลูกค้าก่อน');
    const items = d.rows.filter((r) => r.send > 0).map((r) => ({ order_no: r.order_no, line: r.line, qty: r.send }));
    if (!items.length) throw new Error('ยังไม่ได้ยิงหรือใส่จำนวนที่จะส่งมอบ');
    const pay = $('input[name=dl-pay]:checked').value;
    if (pay === 'โอนเงิน' && !$('#dl-pay-ref').value.trim()) throw new Error('กรุณาใส่รายละเอียดการโอน');
    const body = { by, customer_id: d.customer.id, items, pay_method: pay, pay_ref: $('#dl-pay-ref').value, tax_invoice: $('#dl-tax').checked,
      use_deposit: $('#dl-usedep').checked, note: $('#dl-note').value };
    if (d.debt && d.debt.locked) {
      const ap = await askAdmin('ลูกค้าถูกล็อกการขาย', `ค้างชำระเกินกำหนด ${d.debt.late} วัน ต้องให้แอดมินอนุมัติส่งมอบ`);
      if (!ap) return;
      body.approval = ap.approval;
    }
    const bill = await api('deliver', body);
    toast('บันทึกบิล ' + bill.bill_no + ' แล้ว');
    printDocs(billDoc(bill, 'สำหรับร้าน') + billDoc(bill, 'สำหรับลูกค้า'));
    $('#dl-note').value = ''; $('#dl-pay-ref').value = ''; $('#dl-tax').checked = false;
    pickDeliverCustomer(d.customer);
  });
}

/* ================================================================ หน้าลูกค้า */

async function loadCustomerView(id) {
  const body = $('#cv-body');
  if (!id) { body.innerHTML = ''; S.cv = null; return; }
  body.innerHTML = '<div class="card muted">กำลังโหลด...</div>';
  await busy(null, async () => {
    S.cv = await api('customer', { customer_id: id, days: 120 });
    const fresh = S.cv.customer, i = S.customers.findIndex((c) => c.id === fresh.id);
    if (i >= 0) S.customers[i] = fresh;
    renderCustomerView();
  });
}
function renderCustomerView() {
  const v = S.cv, c = v.customer, d = v.debt, w = v.week, body = $('#cv-body');
  const depLeft = v.deposits.reduce((t, x) => t + x.left, 0);
  const pendQty = v.outstanding.reduce((t, x) => t + x.left, 0);
  body.innerHTML = `
  <div class="card">
    <div class="row wrap"><div class="grow"><h2>${esc(c.display)} ${d.locked ? '<span class="tag dark">ล็อกการขาย</span>' : ''}</h2>
      <div class="muted small">รหัส ${esc(c.code)} · ${esc(v.group.name || c.group)} · เครดิต ${c.credit_days === '' ? num(S.settings.credit_days) : num(c.credit_days)} วัน ${c.phone ? '· โทร ' + esc(c.phone) : ''} ${c.address ? '· ' + esc(c.address) : ''}</div></div>
      <button id="cv-pay" class="primary">รับชำระหนี้</button><button id="cv-stmt">พิมพ์ใบแจ้งหนี้</button><button id="cv-edit" class="ghost">แก้ข้อมูลลูกค้า</button></div>
    <div class="stats">
      <div class="stat ${d.balance > 0 ? 'red' : 'ok'}"><div class="k">ยอดหนี้คงค้าง</div><div class="v">${money(d.balance)}</div></div>
      <div class="stat ${d.overdue > 0 ? 'red' : ''}"><div class="k">เกินกำหนด${d.late ? ` (นานสุด ${d.late} วัน)` : ''}</div><div class="v">${money(d.overdue)}</div></div>
      <div class="stat"><div class="k">มัดจำคงเหลือ</div><div class="v">${money(depLeft)}</div></div>
      <div class="stat"><div class="k">อะไหล่ค้างรับ</div><div class="v">${pendQty} ชิ้น</div></div>
      <div class="stat"><div class="k">ยอดสั่งสัปดาห์นี้ (${esc(thWeek(w.week))})</div><div class="v">${money(w.gross)}</div><div class="small">${esc(w.rate.desc)} ${w.mode === 'discount' ? `<a href="#" id="cv-override">${w.forced ? 'ยกเลิกอนุมัติอัตราสูง' : 'อนุมัติอัตราสูงพิเศษ'}</a>` : ''}</div></div>
    </div>
  </div>
  <div class="grid2">
    <div class="card"><h3>บิลค้างชำระ</h3><div class="table-wrap"><table class="lines"><thead><tr><th>เลขที่</th><th>วันที่</th><th>ครบกำหนด</th><th class="num">ยอดบิล</th><th class="num">คงค้าง</th><th></th></tr></thead><tbody>
      ${d.open.map((x) => `<tr class="${x.late > 0 ? 'bad' : ''}"><td>${x.kind === 'บิล' ? `<a href="#" data-doc="${esc(x.ref)}">${esc(x.ref)}</a>` : esc(x.kind + ' ' + (x.ref || ''))}</td><td>${thDate(x.date)}</td><td>${thDate(x.due_date)}</td>
        <td class="num">${money(x.amount)}</td><td class="num"><b>${money(x.left)}</b></td><td>${x.late > 0 ? `<span class="late">เกิน ${x.late} วัน</span>` : ''}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">ไม่มีหนี้ค้าง</td></tr>'}
      ${d.credit > 0 ? `<tr><td colspan="4">เครดิตคงเหลือ (จ่ายเกิน / ส่วนลดเพิ่ม)</td><td class="num">-${money(d.credit)}</td><td></td></tr>` : ''}
    </tbody></table></div>
    <div class="row gap"><button id="cv-opening" class="ghost mini">+ ยอดหนี้ยกมา</button></div></div>
    <div class="card"><h3>มัดจำ</h3><div class="table-wrap"><table class="lines"><thead><tr><th>ออเดอร์</th><th>วันที่</th><th class="num">มัดจำ</th><th class="num">เหลือ</th><th></th></tr></thead><tbody>
      ${v.deposits.map((x) => `<tr><td><a href="#" data-doc="${esc(x.order_no)}">${esc(x.order_no)}</a></td><td>${thDate(x.date)}</td><td class="num">${money(x.deposit)}</td><td class="num">${money(x.left)}</td>
        <td>${x.status !== 'รอของ' ? `<button class="mini" data-depclose="${esc(x.order_no)}">ปิดยอด</button>` : '<span class="muted small">รอของ</span>'}</td></tr>`).join('') || '<tr><td colspan="5" class="muted">ไม่มี</td></tr>'}
    </tbody></table></div></div>
  </div>
  <div class="card"><div class="row"><h3 class="grow">อะไหล่ค้างรับ</h3><button id="cv-deliver" class="ghost">ไปหน้าส่งมอบ</button><button id="cv-cancel" class="ghost danger">ยกเลิกรายการค้างรับ...</button></div>
    ${pendingTable(v.outstanding, false)}</div>
  <div class="card"><h3>ประวัติ 120 วัน</h3><div class="table-wrap tall"><table class="lines"><thead><tr><th>วันที่</th><th>เอกสาร</th><th>เลขที่</th><th class="num">จำนวนเงิน</th><th>รายละเอียด</th><th>สถานะ</th></tr></thead><tbody>
    ${historyRows(v).map((h) => `<tr class="${h.status === 'ยกเลิก' ? 'void' : ''}"><td>${thDate(h.date)}</td><td>${esc(h.kind)}</td><td><a href="#" data-doc="${esc(h.no)}">${esc(h.no)}</a></td><td class="num">${money(h.amount)}</td><td class="small">${esc(h.info)}</td><td>${esc(h.status)}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">ไม่มี</td></tr>'}
  </tbody></table></div></div>`;
  $('#cv-pay').onclick = () => receivePayment(c, d);
  $('#cv-stmt').onclick = () => printDocs(statementDoc(v));
  $('#cv-edit').onclick = () => editCustomer(c);
  $('#cv-opening').onclick = () => openingBalance(c);
  $('#cv-deliver').onclick = () => { switchTab('deliver'); setPicked('deliver', c); pickDeliverCustomer(c); };
  $('#cv-cancel').onclick = () => cancelPending(v.outstanding);
  const ov = $('#cv-override'); if (ov) ov.onclick = (e) => { e.preventDefault(); overrideWeek(c, w); };
  $$('#cv-body [data-doc]').forEach((a) => { a.onclick = (e) => { e.preventDefault(); openDoc(a.dataset.doc); }; });
  $$('#cv-body [data-depclose]').forEach((b) => { b.onclick = () => closeDeposit(b.dataset.depclose); });
}
function historyRows(v) {
  const out = [];
  v.orders.forEach((o) => out.push({ date: o.date, kind: o.type === URGENT ? 'ออเดอร์ด่วน' : 'ออเดอร์รายสัปดาห์', no: o.order_no, amount: o.gross, info: `โดย ${o.taken_by}${num(o.deposit) ? ' · มัดจำ ' + money(o.deposit) : ''}`, status: o.status }));
  v.bills.forEach((b) => out.push({ date: b.date, kind: 'บิลส่งมอบ', no: b.bill_no, amount: b.net, info: `${b.pay_method}${b.due_date ? ' ครบ ' + thDate(b.due_date) : ''}${truthy(b.tax_invoice) ? ' · ขอใบกำกับ' : ''}`, status: b.status }));
  v.receipts.filter((r) => r.kind === 'รับชำระหนี้').forEach((r) => out.push({ date: r.date, kind: 'รับชำระหนี้', no: r.rv_no, amount: r.amount, info: `${r.method} ${r.ref || ''}`, status: r.status }));
  v.ledger.filter((x) => x.kind === 'ปรับส่วนลด' || x.kind === 'ยกมา').forEach((x) => out.push({ date: x.date, kind: x.kind, no: x.ref, amount: x.amount, info: x.note, status: x.status }));
  return out.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
}
function pendingTable(rows, withCustomer) {
  if (!rows.length) return '<p class="muted">ไม่มีอะไหล่ค้างรับ</p>';
  return `<div class="table-wrap tall"><table class="lines"><thead><tr>${withCustomer ? '<th>ลูกค้า</th>' : ''}<th>ออเดอร์</th><th>วันที่สั่ง</th><th>ประเภท</th><th>รหัส</th><th>ชื่ออะไหล่</th><th class="num">สั่ง</th><th class="num">ส่งแล้ว</th><th class="num">ค้าง</th><th class="num">ค้างกี่วัน</th></tr></thead><tbody>
    ${rows.map((r) => { const age = Math.round((new Date(todayISO()) - new Date(r.date)) / 864e5); return `<tr>${withCustomer ? `<td>${esc(r.customer_name)}</td>` : ''}<td><a href="#" data-doc="${esc(r.order_no)}">${esc(r.order_no)}</a></td><td>${thDate(r.date)}</td><td>${r.type === URGENT ? '<span class="tag red">ด่วน</span>' : 'รายสัปดาห์'}</td>
      <td><b>${esc(r.code)}</b></td><td>${esc(r.name)}</td><td class="num">${r.qty}</td><td class="num">${r.delivered}</td><td class="num"><b>${r.left}</b></td><td class="num ${r.type === URGENT && age > 3 ? 'late' : ''}">${age}</td></tr>`; }).join('')}
  </tbody></table></div>`;
}
function receivePayment(c, d) {
  const box = openModal(`<h2>รับชำระหนี้ · ${esc(c.display)}</h2><p>ยอดหนี้คงค้าง <b>${money(d.balance)}</b> บาท ${d.overdue > 0 ? `(เกินกำหนด <span class="late">${money(d.overdue)}</span>)` : ''}</p>
    <form id="rp-form"><label>จำนวนเงินที่รับ<input name="amount" type="number" step="0.01" min="0.01" value="${d.balance > 0 ? d.balance : ''}" required autofocus></label>
    <label>รับโดย<select name="method"><option>เงินสด</option><option>โอนเงิน</option></select></label>
    <label>รายละเอียด (ธนาคาร / เลขที่เช็ค)<input name="ref"></label><label>หมายเหตุ<input name="note"></label>
    <p class="muted small">ระบบตัดบิลที่เก่าที่สุดก่อน</p>
    <div class="foot"><button type="button" id="rp-cancel">ยกเลิก</button><button class="primary">บันทึก + พิมพ์ใบเสร็จ</button></div></form>`);
  $('#rp-cancel', box).onclick = closeModal;
  $('#rp-form', box).onsubmit = async (ev) => {
    ev.preventDefault(); const f = ev.target;
    await busy($('button.primary', box), async () => {
      const by = staffName();
      const out = await api('receive', { by, customer_id: c.id, amount: num(f.amount.value), method: f.method.value, ref: f.ref.value, note: f.note.value });
      closeModal(); toast('รับชำระแล้ว ' + out.receipt.rv_no);
      printDocs(receiptDoc(out.receipt) + receiptDoc(out.receipt, 'สำหรับร้าน'));
      loadCustomerView(c.id);
    });
  };
}
async function openingBalance(c) {
  const box = openModal(`<h2>ยอดหนี้ยกมา · ${esc(c.display)}</h2><p class="muted">ใช้ใส่หนี้เดิมก่อนเริ่มใช้ระบบ (ใส่ได้หลายรายการ ทีละบิล หรือรวมยอดเดียว)</p>
    <form id="ob-form" class="form2"><label>เลขที่อ้างอิง (บิลเดิม)<input name="ref" value="ยกมา"></label><label>จำนวนเงิน<input name="amount" type="number" step="0.01" required autofocus></label>
    <label>วันที่บิล<input name="date" type="date" value="${todayISO()}"></label><label>ครบกำหนด<input name="due_date" type="date"></label>
    <label class="span2">หมายเหตุ<input name="note"></label>
    <div class="foot span2"><button type="button" id="ob-cancel">ยกเลิก</button><button class="primary">บันทึก</button></div></form>`);
  $('#ob-cancel', box).onclick = closeModal;
  $('#ob-form', box).onsubmit = async (ev) => {
    ev.preventDefault(); const f = ev.target, data = { ref: f.ref.value, amount: num(f.amount.value), date: f.date.value, due_date: f.due_date.value, note: f.note.value };
    let by; try { by = staffName(); } catch (e) { toast(e.message, true); return; }
    const ap = await askAdmin('ใส่ยอดหนี้ยกมา', `${c.display} ${money(data.amount)} บาท`);
    if (!ap) return;
    await busy(null, async () => { await api('opening', Object.assign({ by, customer_id: c.id, approval: ap.approval }, data)); toast('บันทึกยอดยกมาแล้ว'); loadCustomerView(c.id); });
  };
}
async function overrideWeek(c, w) {
  const by = (() => { try { return staffName(); } catch (e) { toast(e.message, true); return ''; } })();
  if (!by) return;
  const ap = await askAdmin(w.forced ? 'ยกเลิกอนุมัติส่วนลดอัตราสูง' : 'อนุมัติส่วนลดอัตราสูงพิเศษ', `${c.display} สัปดาห์ ${thWeek(w.week)} ยอดสั่ง ${money(w.gross)} (ปกติต้องถึง ${money0(w.threshold)}) -> ลด ${w.forced ? w.low : w.high}%`, true);
  if (!ap) return;
  await busy(null, async () => { await api('override', { by, customer_id: c.id, week: w.week, high: !w.forced, note: ap.reason, approval: ap.approval }); toast('บันทึกแล้ว ระบบปรับส่วนลดของที่ส่งมอบไปแล้วให้'); loadCustomerView(c.id); });
}
async function cancelPending(rows) {
  if (!rows.length) { toast('ไม่มีรายการค้างรับ'); return; }
  const box = openModal(`<h2>ยกเลิกรายการค้างรับ</h2><p class="muted">ใช้เมื่อศูนย์ไม่มีของ หรือลูกค้าไม่เอาแล้ว ติ๊กรายการที่จะยกเลิก</p>
    <div class="table-wrap tall"><table class="lines"><tbody>${rows.map((r, i) => `<tr><td><input type="checkbox" data-i="${i}"></td><td>${esc(r.order_no)}</td><td><b>${esc(r.code)}</b></td><td>${esc(r.name)}</td><td class="num">ค้าง ${r.left}</td></tr>`).join('')}</tbody></table></div>
    <div class="foot"><button id="cp-cancel">ปิด</button><button id="cp-ok" class="primary">ยกเลิกรายการที่เลือก</button></div>`, true);
  $('#cp-cancel', box).onclick = closeModal;
  $('#cp-ok', box).onclick = async () => {
    const picked = $$('input[data-i]:checked', box).map((x) => rows[x.dataset.i]);
    if (!picked.length) return;
    let by; try { by = staffName(); } catch (e) { toast(e.message, true); return; }
    closeModal();
    const ap = await askAdmin('ยกเลิกรายการค้างรับ', `${picked.length} รายการ`, true);
    if (!ap) return;
    const byOrder = {};
    picked.forEach((r) => { (byOrder[r.order_no] = byOrder[r.order_no] || []).push(r.line); });
    await busy(null, async () => {
      for (const on of Object.keys(byOrder)) await api('cancelLines', { by, order_no: on, lines: byOrder[on], approval: ap.approval, reason: ap.reason });
      toast('ยกเลิกแล้ว'); if (S.cv) loadCustomerView(S.cv.customer.id);
    });
  };
}
async function closeDeposit(orderNo) {
  const box = openModal(`<h2>ปิดยอดมัดจำ ${esc(orderNo)}</h2><p>ออเดอร์จบแล้วแต่ยังมีมัดจำเหลือ จะทำอย่างไร</p>
    <div class="foot"><button id="dc-x">ปิด</button><button id="dc-refund">คืนเงินลูกค้า</button><button id="dc-credit" class="primary">เก็บเป็นเครดิตหักหนี้</button></div>`);
  $('#dc-x', box).onclick = closeModal;
  const go = async (mode) => {
    let by; try { by = staffName(); } catch (e) { toast(e.message, true); return; }
    closeModal();
    const ap = await askAdmin('ปิดยอดมัดจำ', mode === 'credit' ? 'เก็บเป็นเครดิต' : 'คืนเงินสดให้ลูกค้า');
    if (!ap) return;
    await busy(null, async () => { await api('depositClose', { by, order_no: orderNo, mode, approval: ap.approval }); toast('บันทึกแล้ว'); loadCustomerView(S.cv.customer.id); });
  };
  $('#dc-refund', box).onclick = () => go('refund');
  $('#dc-credit', box).onclick = () => go('credit');
}
function editCustomer(c) {
  c = c || { id: '', code: '', display: '', group: 'GENERAL', credit_days: '', cp_pct: '', cp_roundup: false, active: true };
  const box = openModal(`<h2>${c.id ? 'แก้ข้อมูลลูกค้า' : 'เพิ่มลูกค้าใหม่'}</h2>
    <form id="ce-form" class="form2">
      <label>รหัสลูกค้า${c.id ? '' : ' (เว้นว่าง = ให้ระบบตั้ง)'}<input name="code" value="${esc(c.code)}" ${c.id ? 'disabled' : ''}></label>
      <label>ชื่อที่แสดงบนบิล<input name="display" value="${esc(c.display)}" required autofocus></label>
      <label>คำนำหน้า<input name="title" value="${esc(c.title || '')}"></label><label>ชื่อ<input name="first" value="${esc(c.first || '')}"></label>
      <label>นามสกุล / ชื่อบริษัท<input name="last" value="${esc(c.last || '')}"></label><label>โทร<input name="phone" value="${esc(c.phone || '')}"></label>
      <label class="span2">ที่อยู่<input name="address" value="${esc(c.address || '')}"></label>
      <label>เลขผู้เสียภาษี / บัตรประชาชน<input name="taxid" value="${esc(c.taxid || '')}"></label>
      <label>กลุ่ม<select name="group">${S.groups.map((g) => `<option value="${esc(g.id)}" ${g.id === c.group ? 'selected' : ''}>${esc(g.name || g.id)}</option>`).join('')}</select></label>
      <label>ทุนรวม VAT + % (กลุ่มส่งพม่า)<input name="cp_pct" type="number" step="0.01" value="${esc(c.cp_pct)}"></label>
      <label class="check" style="margin-top:22px"><input type="checkbox" name="cp_roundup" ${truthy(c.cp_roundup) ? 'checked' : ''}> ปัดทุนขึ้นเป็นบาทเต็มก่อนบวก %</label>
      <label>เครดิต (วัน) เว้นว่าง = ค่าเริ่มต้น ${num(S.settings.credit_days || 7)}<input name="credit_days" type="number" value="${esc(c.credit_days)}"></label>
      <label class="check" style="margin-top:22px"><input type="checkbox" name="active" ${c.active === false || c.active === 'FALSE' ? '' : 'checked'}> ใช้งาน</label>
      <label class="span2">หมายเหตุ<input name="note" value="${esc(c.note || '')}"></label>
      <div class="foot span2"><button type="button" id="ce-cancel">ยกเลิก</button><button class="primary">บันทึก</button></div>
    </form>`, true);
  $('#ce-cancel', box).onclick = closeModal;
  $('#ce-form', box).onsubmit = async (ev) => {
    ev.preventDefault(); const f = ev.target;
    const x = { id: c.id, code: f.code.value.trim(), display: f.display.value, title: f.title.value, first: f.first.value, last: f.last.value, phone: f.phone.value,
      address: f.address.value, taxid: f.taxid.value, group: f.group.value, cp_pct: f.cp_pct.value, cp_roundup: f.cp_roundup.checked,
      credit_days: f.credit_days.value, active: f.active.checked, note: f.note.value };
    let by; try { by = staffName(); } catch (e) { toast(e.message, true); return; }
    const ap = await askAdmin('บันทึกข้อมูลลูกค้า', x.display);
    if (!ap) return;
    await busy(null, async () => {
      const saved = await api('saveCustomer', { by, customer: x, approval: ap.approval });
      const i = S.customers.findIndex((y) => y.id === saved.id);
      if (i >= 0) S.customers[i] = saved; else S.customers.push(saved);
      toast('บันทึกลูกค้าแล้ว');
      setPicked('customer', saved); loadCustomerView(saved.id); renderAdminCustomers();
    });
  };
}

/* ================================================================ ค้างรับ / ติดตามหนี้ */

async function loadPending() {
  await busy($('#pd-load'), async () => { S.pending = await api('outstanding', {}); renderPending(); });
}
function renderPending() {
  const q = $('#pd-filter').value.trim().toLowerCase();
  const rows = S.pending.filter((r) => !q || (r.customer_name + ' ' + r.code + ' ' + r.name + ' ' + r.order_no).toLowerCase().includes(q));
  const by = {};
  rows.forEach((r) => { (by[r.customer_id] = by[r.customer_id] || []).push(r); });
  const ids = Object.keys(by).sort((a, b) => by[a][0].customer_name.localeCompare(by[b][0].customer_name, 'th'));
  $('#pd-body').innerHTML = `<div class="card"><div class="stats"><div class="stat"><div class="k">ลูกค้าที่มีของค้าง</div><div class="v">${ids.length}</div></div>
    <div class="stat"><div class="k">รายการ</div><div class="v">${rows.length}</div></div><div class="stat red"><div class="k">ชิ้นที่ค้าง</div><div class="v">${rows.reduce((t, r) => t + r.left, 0)}</div></div></div></div>`
    + ids.map((id) => `<div class="card"><div class="row"><h3 class="grow">${esc(by[id][0].customer_name)} <span class="muted small">${by[id].length} รายการ</span></h3><button class="ghost mini" data-cust="${esc(id)}">ดูลูกค้า</button></div>${pendingTable(by[id], false)}</div>`).join('')
    || '<div class="card muted">ไม่มีอะไหล่ค้างรับ</div>';
  $$('#pd-body [data-doc]').forEach((a) => { a.onclick = (e) => { e.preventDefault(); openDoc(a.dataset.doc); }; });
  $$('#pd-body [data-cust]').forEach((b) => { b.onclick = () => gotoCustomer(b.dataset.cust); });
}
function gotoCustomer(id) { const c = custById(id); if (!c) return; switchTab('customer'); setPicked('customer', c); loadCustomerView(id); }
async function loadDebtors() {
  await busy($('#db-load'), async () => { S.debtors = await api('debtors'); renderDebtors(); });
}
function renderDebtors() {
  const q = $('#db-filter').value.trim().toLowerCase(), late = $('#db-late').checked;
  const rows = S.debtors.filter((r) => (!q || (r.name + ' ' + r.code).toLowerCase().includes(q)) && (!late || r.overdue > 0));
  const sum = (k) => rows.reduce((t, r) => t + num(r[k]), 0);
  $('#db-sum').innerHTML = `<div class="stat red"><div class="k">ยอดหนี้รวม</div><div class="v">${money(sum('balance'))}</div></div>
    <div class="stat red"><div class="k">เกินกำหนดรวม</div><div class="v">${money(sum('overdue'))}</div></div>
    <div class="stat"><div class="k">ลูกค้าที่ค้าง</div><div class="v">${rows.length}</div></div>
    <div class="stat"><div class="k">ถูกล็อกการขาย</div><div class="v">${rows.filter((r) => r.locked).length}</div></div>`;
  $('#db-table tbody').innerHTML = rows.map((r) => `<tr class="clickable ${r.locked ? 'bad' : r.overdue > 0 ? 'warn' : ''}" data-id="${esc(r.customer_id)}"><td>${esc(r.code)}</td><td><b>${esc(r.name)}</b><div class="small muted">${esc(r.phone)}</div></td>
    <td class="num"><b>${money(r.balance)}</b></td><td class="num">${money(r.b0)}</td><td class="num">${r.b1 ? money(r.b1) : ''}</td><td class="num">${r.b2 ? money(r.b2) : ''}</td><td class="num">${r.b3 ? `<span class="late">${money(r.b3)}</span>` : ''}</td>
    <td>${thDate(r.next_due)}</td><td>${r.locked ? '<span class="tag dark">ล็อกการขาย</span>' : r.overdue > 0 ? `<span class="tag red">เกิน ${r.late} วัน</span>` : r.balance < 0 ? '<span class="tag ok">มีเครดิต</span>' : '<span class="tag ok">ปกติ</span>'}</td></tr>`).join('')
    || '<tr><td colspan="9" class="muted">ไม่มีลูกหนี้</td></tr>';
  $$('#db-table tr[data-id]').forEach((tr) => { tr.onclick = () => gotoCustomer(tr.dataset.id); });
}

/* ================================================================ เอกสาร */

async function loadDocs() {
  const from = $('#dc-from').value, to = $('#dc-to').value;
  await busy($('#dc-load'), async () => { S.docs = await api('docs', { from, to }); renderDocs(); });
}
function renderDocs() {
  const d = S.docs;
  const live = (x) => x.status !== 'ยกเลิก';
  const cash = d.receipts.filter(live).filter((r) => r.method === 'เงินสด').reduce((t, r) => t + num(r.amount), 0);
  const tr = d.receipts.filter(live).filter((r) => r.method === 'โอนเงิน').reduce((t, r) => t + num(r.amount), 0);
  const credit = d.bills.filter(live).filter((b) => b.pay_method === 'ค้างจ่าย').reduce((t, b) => t + num(b.due_now), 0);
  $('#dc-body').innerHTML = `<div class="card"><div class="stats">
      <div class="stat"><div class="k">ออเดอร์</div><div class="v">${d.orders.filter(live).length}</div></div>
      <div class="stat"><div class="k">บิลส่งมอบ (สุทธิ)</div><div class="v">${money(d.bills.filter(live).reduce((t, b) => t + num(b.net), 0))}</div></div>
      <div class="stat ok"><div class="k">รับเงินสด</div><div class="v">${money(cash)}</div></div>
      <div class="stat ok"><div class="k">รับเงินโอน</div><div class="v">${money(tr)}</div></div>
      <div class="stat red"><div class="k">ขายเชื่อ (ค้างจ่าย)</div><div class="v">${money(credit)}</div></div></div></div>
    ${docTable('บิลส่งมอบ', d.bills, (b) => [b.bill_no, thDate(b.date) + ' ' + b.time, b.customer_name, money(b.net), b.pay_method + (truthy(b.tax_invoice) ? ' · ใบกำกับ' : ''), b.delivered_by, b.status])}
    ${docTable('ออเดอร์', d.orders, (o) => [o.order_no, thDate(o.date) + ' ' + o.time, o.customer_name, money(o.gross), o.type + (num(o.deposit) ? ' · มัดจำ ' + money(o.deposit) : ''), o.taken_by, o.status])}
    ${docTable('รับเงิน', d.receipts, (r) => [r.rv_no, thDate(r.date) + ' ' + r.time, r.customer_name, money(r.amount), r.kind + ' · ' + r.method + (r.ref ? ' ' + r.ref : ''), r.by, r.status])}`;
  $$('#dc-body [data-doc]').forEach((a) => { a.onclick = (e) => { e.preventDefault(); openDoc(a.dataset.doc); }; });
}
function docTable(title, rows, fn) {
  return `<div class="card"><h3>${esc(title)} (${rows.length})</h3><div class="table-wrap tall"><table class="lines"><thead><tr><th>เลขที่</th><th>วันที่</th><th>ลูกค้า</th><th class="num">จำนวนเงิน</th><th>รายละเอียด</th><th>พนักงาน</th><th>สถานะ</th></tr></thead><tbody>
    ${rows.slice().reverse().map((x) => { const v = fn(x); return `<tr class="${v[6] === 'ยกเลิก' ? 'void' : ''}"><td><a href="#" data-doc="${esc(v[0])}">${esc(v[0])}</a></td>${v.slice(1, 3).map((y) => `<td>${esc(y)}</td>`).join('')}<td class="num">${esc(v[3])}</td>${v.slice(4).map((y) => `<td>${esc(y)}</td>`).join('')}</tr>`; }).join('') || '<tr><td colspan="7" class="muted">ไม่มี</td></tr>'}
  </tbody></table></div></div>`;
}
// เปิดเอกสารจากเลขที่: SO = ออเดอร์, AR = บิลส่งมอบ, RV = ใบเสร็จ
async function openDoc(no) {
  no = String(no || '').trim().toUpperCase();
  if (!no) return;
  await busy(null, async () => {
    let html = '', actions = '';
    if (no.startsWith('SO')) {
      const o = await api('order', { no });
      html = orderDoc(o, 'สำหรับร้าน');
      actions = `<button data-a="print">พิมพ์ใบรับออเดอร์</button>${o.status !== 'ยกเลิก' ? '<button data-a="dep">รับมัดจำเพิ่ม</button><button data-a="void" class="danger">ยกเลิกออเดอร์</button>' : ''}`;
      showDoc(html, actions, {
        print: () => printDocs(orderDoc(o, 'สำหรับร้าน') + orderDoc(o, 'สำหรับลูกค้า')),
        dep: () => addDeposit(o),
        void: async () => { const by = staffName(); const ap = await askAdmin('ยกเลิกออเดอร์ ' + o.order_no, 'มัดจำของออเดอร์นี้จะถือว่าคืนลูกค้าแล้ว', true); if (!ap) return; await api('voidOrder', { by, order_no: o.order_no, approval: ap.approval, reason: ap.reason }); toast('ยกเลิกแล้ว'); openDoc(o.order_no); }
      });
    } else if (no.startsWith('AR')) {
      const b = await api('bill', { no });
      actions = `<button data-a="print">พิมพ์บิล</button>${b.status !== 'ยกเลิก' ? '<button data-a="void" class="danger">ยกเลิกบิล</button>' : ''}`;
      showDoc(billDoc(b, 'สำหรับร้าน'), actions, {
        print: () => printDocs(billDoc(b, 'สำหรับร้าน') + billDoc(b, 'สำหรับลูกค้า')),
        void: async () => { const by = staffName(); const ap = await askAdmin('ยกเลิกบิล ' + b.bill_no, 'อะไหล่จะกลับไปเป็นค้างรับ ยอดหนี้และมัดจำจะคืนกลับ', true); if (!ap) return; await api('voidBill', { by, no: b.bill_no, approval: ap.approval, reason: ap.reason }); toast('ยกเลิกแล้ว'); openDoc(b.bill_no); }
      });
    } else if (no.startsWith('RV')) {
      const r = await api('receipt', { no });
      actions = `<button data-a="print">พิมพ์ใบเสร็จ</button>${r.status !== 'ยกเลิก' ? '<button data-a="void" class="danger">ยกเลิกใบเสร็จ</button>' : ''}`;
      showDoc(receiptDoc(r), actions, {
        print: () => printDocs(receiptDoc(r) + receiptDoc(r, 'สำหรับร้าน')),
        void: async () => { const by = staffName(); const ap = await askAdmin('ยกเลิกใบเสร็จ ' + r.rv_no, '', true); if (!ap) return; await api('voidReceipt', { by, no: r.rv_no, approval: ap.approval, reason: ap.reason }); toast('ยกเลิกแล้ว'); openDoc(r.rv_no); }
      });
    } else throw new Error('เลขที่ต้องขึ้นต้นด้วย SO, AR หรือ RV');
  });
}
function showDoc(html, actions, handlers) {
  const box = openModal(`<div style="overflow:auto">${html}</div><div class="foot">${actions}<button data-a="close">ปิด</button></div>`, true);
  $$('[data-a]', box).forEach((b) => { b.onclick = () => { if (b.dataset.a === 'close') return closeModal(); busy(b, async () => { await handlers[b.dataset.a](); }); }; });
}
function addDeposit(o) {
  const box = openModal(`<h2>รับมัดจำเพิ่ม ${esc(o.order_no)}</h2><form id="ad-form"><label>จำนวนเงิน<input name="amount" type="number" step="0.01" required autofocus></label>
    <label>รับโดย<select name="method"><option>เงินสด</option><option>โอนเงิน</option></select></label><label>รายละเอียด<input name="ref"></label>
    <div class="foot"><button type="button" id="ad-x">ยกเลิก</button><button class="primary">บันทึก</button></div></form>`);
  $('#ad-x', box).onclick = closeModal;
  $('#ad-form', box).onsubmit = async (ev) => {
    ev.preventDefault(); const f = ev.target;
    await busy($('button.primary', box), async () => { const out = await api('addDeposit', { by: staffName(), order_no: o.order_no, amount: num(f.amount.value), method: f.method.value, ref: f.ref.value }); closeModal(); toast('รับมัดจำแล้ว ' + out.rv_no); openDoc(o.order_no); });
  };
}

/* ================================================================ เอกสารพิมพ์ */

function printDocs(html) { $('#print-area').innerHTML = html; setTimeout(() => window.print(), 80); }
function docTop(title, copy, no, extra = '') {
  return `<div class="d-top"><div><div class="d-shop">${esc(S.settings.shop_name || '')}</div><div class="small">สำนักงานใหญ่ · โทร ${esc(S.settings.shop_phone || '')}</div></div>
    <div><div class="d-title">${esc(title)}</div><div class="d-copy">${esc(copy || '')}</div><div style="text-align:right">เลขที่ <b>${esc(no)}</b></div>${extra}</div></div>`;
}
function custBlock(id, name) {
  const c = custById(id) || {};
  return `<div>ลูกค้า <b>${esc(name || c.display || '')}</b> (รหัส ${esc(c.code || '')})</div>
    ${custFull(c) || c.address ? `<div class="small">${esc(custFull(c))} ${esc(c.address || '')}</div>` : ''}${c.phone ? `<div class="small">โทร ${esc(c.phone)}</div>` : ''}`;
}
function orderDoc(o, copy) {
  const isW = o.type === WEEKLY, lines = o.lines || [];
  const est = lines.reduce((t, l) => t + num(l.est_net), 0);
  const rows = lines.map((l, i) => `<tr class="${num(l.cancelled) ? 'void' : ''}"><td class="num">${i + 1}</td><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="num">${l.qty}</td>
    <td class="num">${money(l.price)}</td><td class="num">${money(num(l.qty) * num(l.price))}</td><td class="num">${money(l.est_net)}</td></tr>`).join('');
  const dep = num(o.deposit);
  return `<div class="doc">${docTop('ใบรับออเดอร์', copy, o.order_no, `<div style="text-align:right"><span class="stamp">${isW ? 'สั่งรายสัปดาห์' : 'สั่งด่วน 1-3 วัน'}</span></div>`)}
    <div class="d-head"><div>${custBlock(o.customer_id, o.customer_name)}</div>
      <div>วันที่รับออเดอร์ <b>${thDate(o.date)}</b> เวลา ${esc(o.time)}<br>${isW ? `รอบสัปดาห์ ${esc(thWeek(o.week))}<br>` : ''}ผู้รับออเดอร์ <b>${esc(o.taken_by)}</b>${o.status === 'ยกเลิก' ? '<br><b style="color:#c62828">ยกเลิกแล้ว</b>' : ''}</div></div>
    <table><thead><tr><th style="width:5%">ที่</th><th style="width:16%">รหัสอะไหล่</th><th>ชื่ออะไหล่</th><th style="width:13%">รุ่น</th><th style="width:7%">จำนวน</th><th style="width:10%">ราคาขาย</th><th style="width:11%">รวมเงิน</th><th style="width:11%">สุทธิ*</th></tr></thead><tbody>${rows}</tbody></table>
    <div class="d-sum"><div>(${esc(bahtText(est))})</div><div>รวมราคาขาย</div><div class="num">${money(o.gross)}</div>
      <div>${esc(o.rate ? o.rate.desc : '')}</div><div>สุทธิโดยประมาณ*</div><div class="num big">${money(est)}</div>
      ${dep ? `<div></div><div>วางมัดจำ (${esc(o.deposit_method)})</div><div class="num">${money(dep)}</div><div></div><div>คงเหลือจ่าย*</div><div class="num big">${money(est - dep)}</div>` : ''}</div>
    <div class="d-note">* ${isW ? esc(S.settings.order_note || '') : 'ราคาสุทธิคิดตอนส่งมอบ'} ${o.note ? '<br>หมายเหตุ: ' + esc(o.note) : ''}</div>
    <div class="d-sign"><div>ลงชื่อ...........................................<br>ผู้รับออเดอร์ (${esc(o.taken_by)})</div><div>ลงชื่อ...........................................<br>ลูกค้า</div></div></div>`;
}
function billDoc(b, copy) {
  const lines = b.lines || [], credit = b.pay_method === 'ค้างจ่าย';
  const rows = lines.map((l, i) => `<tr><td class="num">${i + 1}</td><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="num">${l.qty}</td><td class="num">${money(l.price)}</td>
    <td class="num">${money(l.unit_net)}</td><td class="num">${money(l.net)}</td><td class="small">${esc(l.order_no)}</td></tr>`).join('');
  const pay = credit ? `<b>ค้างจ่าย</b> ครบกำหนดชำระ <b>${thDate(b.due_date)}</b>` : `ชำระโดย <b>${esc(b.pay_method)}</b> ${esc(b.pay_ref || '')}`;
  return `<div class="doc">${docTop('บิลส่งมอบอะไหล่', copy, b.bill_no, truthy(b.tax_invoice) ? '<div style="text-align:right" class="small">ขอใบกำกับภาษีฉบับเต็ม</div>' : '')}
    <div class="d-head"><div>${custBlock(b.customer_id, b.customer_name)}</div><div>วันที่ <b>${thDate(b.date)}</b> เวลา ${esc(b.time)}<br>ผู้ส่งมอบ <b>${esc(b.delivered_by)}</b>${b.status === 'ยกเลิก' ? '<br><b style="color:#c62828">ยกเลิกแล้ว</b>' : ''}</div></div>
    <table><thead><tr><th style="width:5%">ที่</th><th style="width:16%">รหัสอะไหล่</th><th>ชื่ออะไหล่</th><th style="width:12%">รุ่น</th><th style="width:7%">จำนวน</th><th style="width:10%">ราคาขาย</th><th style="width:10%">สุทธิ/ชิ้น</th><th style="width:11%">จำนวนเงิน</th><th style="width:11%">ออเดอร์</th></tr></thead><tbody>${rows}</tbody></table>
    <div class="d-sum"><div>(${esc(bahtText(b.due_now))})</div><div>รวมราคาขาย</div><div class="num">${money(b.gross)}</div>
      <div>${pay}</div><div>ส่วนลด</div><div class="num">${money(b.discount)}</div>
      <div></div><div>สุทธิ</div><div class="num">${money(b.net)}</div>
      ${num(b.deposit_used) ? `<div></div><div>หักเงินมัดจำ</div><div class="num">${money(b.deposit_used)}</div>` : ''}
      <div></div><div class="big">${credit ? 'ยอดค้างชำระ' : 'ยอดชำระ'}</div><div class="num big">${money(b.due_now)}</div></div>
    ${(b.adjust || []).length ? `<div class="d-note">ปรับส่วนลดสัปดาห์ภายหลัง: ${b.adjust.map((a) => money(a.amount)).join(', ')} บาท (หักในบัญชีลูกหนี้)</div>` : ''}
    <div class="d-note">${b.note ? 'หมายเหตุ: ' + esc(b.note) + '<br>' : ''}${esc(S.settings.bill_footer || '')}</div>
    <div class="d-sign"><div>ลงชื่อ...........................................<br>ผู้รับสินค้า</div><div>ลงชื่อ...........................................<br>ผู้ส่งมอบ (${esc(b.delivered_by)})</div><div>ลงชื่อ...........................................<br>ผู้ตรวจสอบ</div></div></div>`;
}
function receiptDoc(r, copy = 'สำหรับลูกค้า') {
  const applied = String(r.applied || '').split(',').filter(Boolean).map((x) => { const [k, v] = x.split(':'); return `<tr><td>${esc(k)}</td><td class="num">${money(v)}</td></tr>`; }).join('');
  const title = r.kind === 'มัดจำ' ? 'ใบรับเงินมัดจำ' : 'ใบเสร็จรับเงิน';
  return `<div class="doc half">${docTop(title, copy, r.rv_no)}
    <div class="d-head"><div>${custBlock(r.customer_id, r.customer_name)}</div><div>วันที่ <b>${thDate(r.date)}</b> เวลา ${esc(r.time)}<br>ผู้รับเงิน <b>${esc(r.by)}</b>${r.status === 'ยกเลิก' ? '<br><b style="color:#c62828">ยกเลิกแล้ว</b>' : ''}</div></div>
    <p style="font-size:16px">ได้รับเงินจำนวน <b>${money(r.amount)}</b> บาท (${esc(bahtText(r.amount))})<br>โดย <b>${esc(r.method)}</b> ${esc(r.ref || '')} ${r.order_no ? '· ออเดอร์ ' + esc(r.order_no) : ''}</p>
    ${applied ? `<table style="width:60%"><thead><tr><th>ชำระบิล</th><th class="num">จำนวนเงิน</th></tr></thead><tbody>${applied}</tbody></table>` : ''}
    ${r.kind === 'รับชำระหนี้' ? `<p>ยอดหนี้คงเหลือหลังรับชำระ <b>${money(r.debt_after)}</b> บาท</p>` : ''}
    <div class="d-sign"><div>ลงชื่อ...........................................<br>ผู้จ่ายเงิน</div><div>ลงชื่อ...........................................<br>ผู้รับเงิน (${esc(r.by)})</div></div></div>`;
}
function statementDoc(v) {
  const c = v.customer, d = v.debt;
  const rows = d.open.map((x, i) => `<tr><td class="num">${i + 1}</td><td>${esc(x.kind === 'บิล' ? x.ref : x.kind + ' ' + (x.ref || ''))}</td><td>${thDate(x.date)}</td><td>${thDate(x.due_date)}</td><td class="num">${money(x.amount)}</td><td class="num">${money(x.left)}</td><td>${x.late > 0 ? 'เกิน ' + x.late + ' วัน' : ''}</td></tr>`).join('');
  return `<div class="doc">${docTop('ใบแจ้งหนี้ / ใบวางบิล', 'สำหรับลูกค้า', c.code + '-' + todayISO().replace(/-/g, ''))}
    <div class="d-head"><div>${custBlock(c.id, c.display)}</div><div>วันที่ <b>${thDate(todayISO())}</b></div></div>
    <table><thead><tr><th style="width:5%">ที่</th><th>เลขที่บิล</th><th>วันที่</th><th>ครบกำหนด</th><th class="num">ยอดบิล</th><th class="num">คงค้าง</th><th>หมายเหตุ</th></tr></thead><tbody>${rows || '<tr><td colspan="7">ไม่มียอดค้าง</td></tr>'}</tbody></table>
    <div class="d-sum">${d.credit > 0 ? `<div></div><div>หักเครดิตคงเหลือ</div><div class="num">${money(d.credit)}</div>` : ''}
      <div>(${esc(bahtText(d.balance))})</div><div class="big">ยอดที่ต้องชำระ</div><div class="num big">${money(d.balance)}</div></div>
    <div class="d-sign"><div>ลงชื่อ...........................................<br>ผู้วางบิล</div><div>ลงชื่อ...........................................<br>ผู้รับวางบิล วันที่ ........../........../..........</div></div></div>`;
}

/* ================================================================ Excel */

function saveXlsx(name, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(([title, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), title.slice(0, 31)));
  XLSX.writeFile(wb, name);
}
async function exportStock() {
  const from = $('#dc-from').value, to = $('#dc-to').value;
  await busy($('#dc-x-stock'), async () => {
    const d = await api('exportLines', { from, to });
    const bills = {}; d.bills.forEach((b) => { bills[b.bill_no] = b; });
    // คอลัมน์แบบ "บันทึกขายอะไหล่" เดิม: รวมเงิน/ส่วนลด/เหลือ อยู่บรรทัดแรกของบิล
    const rows = [['ลำดับ', 'ใบเสร็จเลขที่', 'วันที่', 'รหัสอะไหล่', 'ชื่ออะไหล่', 'รุ่น', 'จำนวน', 'หน่วยละ', 'จำนวนเงิน', 'ชื่อผู้ซื้อ', 'รวมเงิน', 'ส่วนลด', 'เหลือ', 'ราคาสุทธิ', 'ออเดอร์', 'ชำระโดย', 'หักมัดจำ']];
    let n = 0, last = '';
    d.lines.filter((l) => l.status !== 'ยกเลิก').sort((a, b) => (a.bill_no < b.bill_no ? -1 : a.bill_no > b.bill_no ? 1 : a.line - b.line)).forEach((l) => {
      const b = bills[l.bill_no], first = l.bill_no !== last; last = l.bill_no;
      rows.push([++n, l.bill_no, thDate(l.date), l.code, l.name, l.model, num(l.qty), num(l.price), num(l.gross), first ? b.customer_name : '', first ? num(b.gross) : '', first ? num(b.discount) : '', first ? num(b.net) : '', num(l.net), l.order_no, first ? b.pay_method : '', first ? num(b.deposit_used) : '']);
    });
    saveXlsx(`บันทึกขายอะไหล่-ขายส่ง ${from}_${to}.xlsx`, [['บันทึกขายอะไหล่', rows]]);
  });
}
async function exportTax() {
  const from = $('#dc-from').value, to = $('#dc-to').value;
  await busy($('#dc-x-tax'), async () => {
    const d = await api('exportLines', { from, to });
    const rows = [['เลขที่บิล', 'วันที่', 'รหัสลูกค้า', 'ชื่อบนบิล', 'ชื่อ-นามสกุล / บริษัท', 'ที่อยู่', 'เลขผู้เสียภาษี', 'มูลค่าก่อน VAT', 'VAT 7%', 'รวม', 'ชำระโดย']];
    d.bills.filter((b) => b.status !== 'ยกเลิก' && truthy(b.tax_invoice)).forEach((b) => {
      const c = custById(b.customer_id) || {}, net = num(b.net), ex = round2(net / 1.07);
      rows.push([b.bill_no, thDate(b.date), c.code || '', b.customer_name, custFull(c), c.address || '', c.taxid || '', ex, round2(net - ex), net, b.pay_method]);
    });
    const lines = [['เลขที่บิล', 'รหัสอะไหล่', 'ชื่ออะไหล่', 'จำนวน', 'ราคาต่อหน่วย (รวม VAT)', 'จำนวนเงิน (รวม VAT)']];
    const want = new Set(rows.slice(1).map((r) => r[0]));
    d.lines.filter((l) => want.has(l.bill_no) && l.status !== 'ยกเลิก').forEach((l) => lines.push([l.bill_no, l.code, l.name, num(l.qty), num(l.unit_net), num(l.net)]));
    saveXlsx(`ข้อมูลใบกำกับภาษี ${from}_${to}.xlsx`, [['ใบกำกับภาษี', rows], ['รายการ', lines]]);
  });
}
async function exportMoney() {
  if (!S.docs) await loadDocs();
  const d = S.docs, rows = [['เลขที่', 'วันที่', 'เวลา', 'ลูกค้า', 'ประเภท', 'วิธีรับเงิน', 'รายละเอียด', 'จำนวนเงิน', 'ผู้รับเงิน', 'สถานะ', 'ชำระบิล']];
  d.receipts.forEach((r) => rows.push([r.rv_no, thDate(r.date), r.time, r.customer_name, r.kind, r.method, r.ref, num(r.amount), r.by, r.status, r.applied]));
  saveXlsx(`รับเงินประจำวัน ${$('#dc-from').value}_${$('#dc-to').value}.xlsx`, [['รับเงิน', rows]]);
}
async function exportOrders() {
  if (!S.docs) await loadDocs();
  const rows = [['เลขที่', 'วันที่', 'เวลา', 'ลูกค้า', 'ประเภท', 'รอบสัปดาห์', 'รวมราคาขาย', 'มัดจำ', 'สถานะ', 'ผู้รับออเดอร์', 'หมายเหตุ']];
  S.docs.orders.forEach((o) => rows.push([o.order_no, thDate(o.date), o.time, o.customer_name, o.type, o.type === WEEKLY ? thWeek(o.week) : '', num(o.gross), num(o.deposit), o.status, o.taken_by, o.note]));
  saveXlsx(`ออเดอร์ ${$('#dc-from').value}_${$('#dc-to').value}.xlsx`, [['ออเดอร์', rows]]);
}
function exportPending() {
  const rows = [['ลูกค้า', 'ออเดอร์', 'วันที่สั่ง', 'ประเภท', 'รหัสอะไหล่', 'ชื่ออะไหล่', 'รุ่น', 'สั่ง', 'ส่งแล้ว', 'ยกเลิก', 'ค้าง']];
  S.pending.forEach((r) => rows.push([r.customer_name, r.order_no, thDate(r.date), r.type, r.code, r.name, r.model, r.qty, r.delivered, r.cancelled, r.left]));
  const sum = {}; S.pending.forEach((r) => { sum[r.code] = sum[r.code] || [r.code, r.name, r.model, 0]; sum[r.code][3] += r.left; });
  saveXlsx(`อะไหล่ค้างรับ ${todayISO()}.xlsx`, [['ค้างรับรายลูกค้า', rows], ['รวมตามรหัส (ใช้สั่งของ)', [['รหัสอะไหล่', 'ชื่ออะไหล่', 'รุ่น', 'จำนวนค้างรวม'], ...Object.values(sum)]]]);
}
function exportDebtors() {
  const rows = [['รหัส', 'ลูกค้า', 'โทร', 'ยอดหนี้', 'ยังไม่ถึงกำหนด', 'เกิน 1-7 วัน', 'เกิน 8-30 วัน', 'เกิน 30 วัน', 'เกินนานสุด (วัน)', 'ครบกำหนดถัดไป', 'ล็อกการขาย']];
  S.debtors.forEach((r) => rows.push([r.code, r.name, r.phone, r.balance, r.b0, r.b1, r.b2, r.b3, r.late, thDate(r.next_due), r.locked ? 'ล็อก' : '']));
  saveXlsx(`ลูกหนี้ ${todayISO()}.xlsx`, [['ลูกหนี้', rows]]);
}

/* ================================================================ แอดมิน */

function renderAdmin() {
  const f = $('#ad-settings');
  ['shop_name', 'shop_phone', 'retail_sheet_id', 'urgent_pct', 'week_cutoff', 'credit_days', 'lock_days', 'order_note', 'bill_footer', 'gemini_model'].forEach((k) => { f[k].value = S.settings[k] == null ? '' : S.settings[k]; });
  f.new_admin_code.value = '';
  $('#ad-gemini').textContent = S.gemini ? 'อ่านรูป: ใส่รหัส Gemini แล้ว' : 'อ่านรูป: ยังไม่ได้ใส่รหัส Gemini';
  $('#ad-groups tbody').innerHTML = S.groups.map((g, i) => `<tr><td><input data-i="${i}" data-k="id" value="${esc(g.id)}" ${g.id === 'GENERAL' ? 'disabled' : ''} style="width:90px"></td><td><input data-i="${i}" data-k="name" value="${esc(g.name)}" style="min-width:200px"></td>
    <td><select data-i="${i}" data-k="mode"><option value="discount" ${g.mode !== 'costplus' ? 'selected' : ''}>ลด % จากราคาขาย</option><option value="costplus" ${g.mode === 'costplus' ? 'selected' : ''}>ทุนรวม VAT + %</option></select></td>
    <td><input data-i="${i}" data-k="low_pct" type="number" step="0.01" value="${esc(g.low_pct)}" class="qty"></td><td><input data-i="${i}" data-k="high_pct" type="number" step="0.01" value="${esc(g.high_pct)}" class="qty"></td>
    <td><input data-i="${i}" data-k="threshold" type="number" value="${esc(g.threshold)}" style="width:100px"></td></tr>`).join('');
  $$('#ad-groups [data-k]').forEach((x) => { x.onchange = () => { S.groups[x.dataset.i][x.dataset.k] = x.value; }; });
  const sel = $('#ad-cgroup'), cur = sel.value;
  sel.innerHTML = '<option value="">ทุกกลุ่ม</option>' + S.groups.map((g) => `<option value="${esc(g.id)}">${esc(g.name || g.id)}</option>`).join('');
  sel.value = cur;
  renderAdminCustomers();
}
function renderAdminCustomers() {
  const q = $('#ad-cfilter').value.trim().toLowerCase(), g = $('#ad-cgroup').value;
  const rows = S.customers.filter((c) => (!g || c.group === g) && (!q || (c.code + ' ' + c.display + ' ' + c.phone).toLowerCase().includes(q)));
  $('#ad-cust tbody').innerHTML = rows.slice(0, 500).map((c) => `<tr class="${c.active === false || c.active === 'FALSE' ? 'void' : ''}"><td>${esc(c.code)}</td><td>${esc(c.display)}</td><td>${esc((groupById(c.group) || {}).name || c.group)}</td>
    <td class="num">${c.cp_pct === '' ? '' : esc(c.cp_pct) + '%' + (truthy(c.cp_roundup) ? ' ปัดขึ้น' : '')}</td><td class="num">${esc(c.credit_days)}</td><td>${esc(c.phone)}</td><td><button class="mini" data-id="${esc(c.id)}">แก้</button></td></tr>`).join('');
  $$('#ad-cust [data-id]').forEach((b) => { b.onclick = () => editCustomer(custById(b.dataset.id)); });
}
async function saveGroups(btn) {
  let by; try { by = staffName(); } catch (e) { toast(e.message, true); return; }
  const ap = await askAdmin('บันทึกกลุ่มและส่วนลด');
  if (!ap) return;
  await busy(btn, async () => { S.groups = await api('saveGroups', { by, groups: S.groups, approval: ap.approval }); toast('บันทึกกลุ่มแล้ว'); renderAdmin(); });
}
async function saveSettings(ev) {
  ev.preventDefault();
  const f = ev.target, set = {};
  ['shop_name', 'shop_phone', 'retail_sheet_id', 'urgent_pct', 'week_cutoff', 'credit_days', 'lock_days', 'order_note', 'bill_footer', 'gemini_model'].forEach((k) => { set[k] = f[k].value.trim(); });
  let by; try { by = staffName(); } catch (e) { toast(e.message, true); return; }
  const ap = await askAdmin('บันทึกการตั้งค่า');
  if (!ap) return;
  await busy($('button[type=submit]', f), async () => {
    const changedRetail = set.retail_sheet_id !== String(S.settings.retail_sheet_id || '');
    S.settings = await api('saveSettings', { by, settings: set, new_admin_code: f.new_admin_code.value, approval: ap.approval });
    toast('บันทึกการตั้งค่าแล้ว');
    if (f.new_admin_code.value) toast('เปลี่ยนรหัสแอดมินแล้ว ใช้รหัสใหม่ในการอนุมัติครั้งต่อไป');
    if (changedRetail) syncParts(true);
    renderAdmin();
  });
}
async function loadAudit() {
  await busy($('#ad-audit'), async () => {
    const rows = await api('audit', { limit: 300 });
    $('#ad-audit-table tbody').innerHTML = rows.map((a) => `<tr><td class="small">${esc(String(a.ts).replace('T', ' '))}</td><td>${esc(a.by)}</td><td>${esc(a.action)}</td><td>${esc(a.ref)}</td><td>${esc(a.approved)}</td><td class="small muted">${esc(String(a.detail).slice(0, 160))}</td></tr>`).join('');
  });
}

/* ================================================================ ผูกปุ่ม */

const TAB_LOAD = { pending: loadPending, debt: loadDebtors, admin: renderAdmin, docs: () => { if (!S.docs) loadDocs(); } };
function switchTab(name) {
  $$('.tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  $$('.tab').forEach((t) => t.classList.toggle('hidden', t.id !== 'tab-' + name));
  if (TAB_LOAD[name]) TAB_LOAD[name]();
  if (name === 'order') $('#order-scan').focus();
  if (name === 'deliver') $('#deliver-scan').focus();
}
function bindStatic() {
  $('#setup-save').onclick = () => {
    const u = $('#setup-url').value.trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(u) && !/^http:\/\/localhost/.test(u)) { $('#setup-err').textContent = 'ลิงก์ต้องขึ้นต้น https://script.google.com/ และลงท้าย /exec'; return; }
    LS.set('apiUrl', u); S.apiUrl = u; show('scr-unlock');
  };
  $('#unlock-reset-url').onclick = (e) => { e.preventDefault(); LS.del('apiUrl'); S.apiUrl = ''; show('scr-setup'); };
  $('#unlock-form').onsubmit = async (ev) => {
    ev.preventDefault(); $('#unlock-err').textContent = '';
    const user = $('#login-user').value.trim();
    try {
      const r = await api('login', { username: user, pin: $('#unlock-code').value });
      S.token = r.token; S.user = r.user; LS.set('token', r.token); LS.set('user', JSON.stringify(r.user)); LS.set('lastUser', user);
      $('#unlock-code').value = ''; await start();
    }
    catch (e) { $('#unlock-err').textContent = e.message; }
  };
  $('#logout').onclick = (e) => { e.preventDefault(); if (confirm('ออกจากระบบ?')) logout(); };
  $$('.tabs button').forEach((b) => { b.onclick = () => switchTab(b.dataset.tab); });
  $('#modal').onclick = (e) => { if (e.target.id === 'modal') closeModal(); };
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('#modal').classList.contains('hidden')) closeModal(); });

  // ออเดอร์
  $$('#order-type button').forEach((b) => { b.onclick = () => { S.order.type = b.dataset.v; $$('#order-type button').forEach((x) => x.classList.toggle('on', x === b)); renderOrder(); }; });
  $('#order-scan').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); orderScan(e.target.value); e.target.value = ''; } };
  $('#order-search').onclick = () => openPartSearch($('#order-scan').value, (p) => addOrderPart(p, Math.max(1, num($('#order-qty').value))));
  $('#order-manual').onclick = manualLine;
  $('#order-file').onchange = (e) => { const f = Array.from(e.target.files); e.target.value = ''; readFiles(f); };
  $('#order-deposit').oninput = () => { $('#ot-dep').textContent = money($('#order-deposit').value); };
  $$('.pct-btns button').forEach((b) => { b.onclick = (e) => { e.preventDefault(); $('#order-deposit').value = Math.ceil(num($('#ot-net').textContent.replace(/,/g, '')) * num(b.dataset.pct) / 100 / 100) * 100; $('#order-deposit').oninput(); }; });
  $('#order-clear').onclick = () => { const c = S.order.customer; newOrder(); S.order.customer = c; renderOrder(); };
  $('#order-save').onclick = (e) => saveOrder(e.currentTarget);

  // ส่งมอบ
  $('#deliver-scan').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); deliverScan(e.target.value); e.target.value = ''; } };
  $('#deliver-all').onclick = () => { if (S.dl) { S.dl.rows.forEach((r) => { r.send = r.left; }); renderDeliver(); } };
  $('#deliver-none').onclick = () => { if (S.dl) { S.dl.rows.forEach((r) => { r.send = 0; }); renderDeliver(); } };
  $$('input[name=dl-pay]').forEach((r) => { r.onchange = () => $('#dl-pay-ref').classList.toggle('hidden', $('input[name=dl-pay]:checked').value !== 'โอนเงิน'); });
  $('#dl-usedep').onchange = deliverTotals;
  $('#deliver-save').onclick = (e) => saveDelivery(e.currentTarget);

  // ลูกค้า ค้างรับ หนี้ เอกสาร
  $('#cv-refresh').onclick = () => { if (S.cv) loadCustomerView(S.cv.customer.id); };
  $('#cv-new').onclick = () => editCustomer(null);
  $('#pd-load').onclick = loadPending; $('#pd-filter').oninput = renderPending; $('#pd-excel').onclick = exportPending;
  $('#db-load').onclick = loadDebtors; $('#db-filter').oninput = renderDebtors; $('#db-late').onchange = renderDebtors; $('#db-excel').onclick = exportDebtors;
  $('#dc-load').onclick = loadDocs;
  $('#dc-open').onclick = () => openDoc($('#dc-no').value);
  $('#dc-no').onkeydown = (e) => { if (e.key === 'Enter') openDoc(e.target.value); };
  $('#dc-x-stock').onclick = exportStock; $('#dc-x-tax').onclick = exportTax;
  $('#dc-x-money').onclick = () => busy($('#dc-x-money'), exportMoney); $('#dc-x-orders').onclick = () => busy($('#dc-x-orders'), exportOrders);

  // แอดมิน
  $('#ad-group-add').onclick = () => { S.groups.push({ id: 'NEW' + S.groups.length, name: 'กลุ่มใหม่', mode: 'discount', low_pct: 25, high_pct: 29, threshold: 50000 }); renderAdmin(); };
  $('#ad-group-save').onclick = (e) => saveGroups(e.currentTarget);
  $('#ad-settings').onsubmit = saveSettings;
  $('#ad-cfilter').oninput = renderAdminCustomers; $('#ad-cgroup').onchange = renderAdminCustomers;
  $('#ad-cexcel').onclick = () => saveXlsx(`ลูกค้าขายส่ง ${todayISO()}.xlsx`, [['ลูกค้า', [['รหัส', 'ชื่อบนบิล', 'คำนำหน้า', 'ชื่อ', 'นามสกุล', 'ที่อยู่', 'โทร', 'เลขผู้เสียภาษี', 'กลุ่ม', 'ทุน+%', 'ปัดขึ้น', 'เครดิต(วัน)'],
    ...S.customers.map((c) => [c.code, c.display, c.title, c.first, c.last, c.address, c.phone, c.taxid, c.group, c.cp_pct, truthy(c.cp_roundup) ? 'ใช่' : '', c.credit_days])]]]);
  $('#ad-audit').onclick = loadAudit;
  $('#ad-reparts').onclick = (e) => { e.preventDefault(); syncParts(true); };
}

if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
window.addEventListener('DOMContentLoaded', boot);
