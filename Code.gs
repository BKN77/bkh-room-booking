/** ระบบจองห้องพักเวร รพ.บึงกาฬ — Backend (Google Apps Script + Google Sheets)
 *  ขั้นตอน: Extensions > Apps Script ในชีต > วางไฟล์นี้ + Index.html > รัน setup() 1 ครั้ง > Deploy เป็น Web app */
/** ถ้าสร้าง Apps Script แยกจากชีต (ไม่ได้เปิดจากเมนู ส่วนขยาย ในชีต) ให้ใส่ ID ของชีตตรงนี้ (ข้อความยาวๆ ใน URL ระหว่าง /d/ และ /edit) */
const SHEET_ID = '136_-qaN642MiFz55VaLuZkZojdWmgwZItzrzGuMij_Q';
const ADMIN_U = 'BKH77', ADMIN_P = 'BKH11040', SALT = 'bkh|', TZ = 'Asia/Bangkok';
const H = {
  Users: ['username', 'pass', 'name', 'dept', 'phone', 'status', 'createdAt'],
  Rooms: ['id', 'name', 'note', 'active'],
  Bookings: ['tx', 'createdAt', 'roomId', 'roomName', 'inspect', 'roomStatus', 'name', 'username', 'phone', 'shift', 'channel', 'cin', 'cout', 'nights', 'ciStatus', 'coStatus', 'note', 'status'],
  Inspections: ['id', 'date', 'roomId', 'inspector', 'result', 'tx', 'note'],
  Config: ['key', 'value']
};
const DEF = {
  shifts: ['เวรเช้า', 'เวรบ่าย', 'เวรดึก', 'เวรเช้า-บ่าย', 'เวรบ่าย-ดึก', 'ควบเวร'],
  channels: ['เว็บไซต์', 'โทรศัพท์', 'LINE', 'Walk-in'], maxNights: 7, seq: 0
};

function setup() {
  if (/^AKfycb/.test(SHEET_ID)) throw new Error('SHEET_ID ผิด: นี่คือ Deployment ID ไม่ใช่รหัสของ Google Sheet (ดูรหัสชีตจาก URL ระหว่าง /d/ และ /edit)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
  if (!ss) throw new Error('ไม่พบชีต: ให้ใส่ SHEET_ID ด้านบนของไฟล์');
  PropertiesService.getScriptProperties().setProperty('SID', ss.getId());
  Object.keys(H).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    s.getRange(1, 1, s.getMaxRows(), H[n].length).setNumberFormat('@');
    s.getRange(1, 1, 1, H[n].length).setValues([H[n]]).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  if (all_('Rooms').length === 0) [1, 2, 3, 4].forEach(i => add_('Rooms', { id: 'R' + i, name: 'ห้องพักเวร ' + i, note: '', active: 'TRUE' }));
  const c = cfg_();
  Object.keys(DEF).forEach(k => { if (c[k] === undefined) setCfg_(k, DEF[k]); });
  all_('Bookings').forEach(b => { if (b.roomStatus === 'ว่าง') upd_('Bookings', 'tx', b.tx, { roomStatus: 'เช็คเอาท์แล้ว' }); }); // แปลงข้อมูลเก่า
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('ระบบจองห้องพักเวร โรงพยาบาลบึงกาฬ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** รับคำสั่งจากหน้าเว็บที่โฮสต์ภายนอก (เช่น GitHub Pages) */
function doPost(e) {
  let r;
  try { const q = JSON.parse(e.postData.contents); r = api(q.a, q.t, q.p); }
  catch (x) { r = { ok: false, m: String(x.message || x) }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */
function sh_(n) {
  let id = SHEET_ID || PropertiesService.getScriptProperties().getProperty('SID');
  if (!id) { try { const a = SpreadsheetApp.getActive(); if (a) id = a.getId(); } catch (e) { } }
  if (!id) fail_('ยังไม่ได้ตั้งค่า: กรุณารันฟังก์ชัน setup() หรือใส่ SHEET_ID ใน Code.gs แล้ว Deploy ใหม่');
  if (/^AKfycb/.test(id)) fail_('SHEET_ID ผิด: ค่านี้คือรหัสการทำให้ใช้งานได้ (Deployment ID) ไม่ใช่รหัสของ Google Sheet ให้ใช้ข้อความระหว่าง /d/ และ /edit ใน URL ของชีต แล้วรัน setup ใหม่');
  const s = SpreadsheetApp.openById(id).getSheetByName(n);
  if (!s) fail_('ไม่พบชีต ' + n + ' กรุณารันฟังก์ชัน setup()');
  return s;
}
function all_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], TZ, 'yyyy-MM-dd') : String(r[i])); return o; });
}
function add_(n, o) { sh_(n).appendRow(H[n].map(k => String(o[k] === undefined ? '' : o[k]))); }
function upd_(n, kc, key, patch) {
  const s = sh_(n), v = s.getDataRange().getValues(), h = v[0], ci = h.indexOf(kc);
  for (let i = 1; i < v.length; i++) if (String(v[i][ci]) === String(key)) {
    Object.keys(patch).forEach(k => { const j = h.indexOf(k); if (j >= 0) s.getRange(i + 1, j + 1).setValue(String(patch[k])); });
    return true;
  }
  return false;
}
function del_(n, kc, key) {
  const s = sh_(n), v = s.getDataRange().getValues(), ci = v[0].indexOf(kc);
  for (let i = 1; i < v.length; i++) if (String(v[i][ci]) === String(key)) { s.deleteRow(i + 1); return true; }
  return false;
}
function cfg_() { const o = {}; all_('Config').forEach(r => { try { o[r.key] = JSON.parse(r.value); } catch (e) { } }); return o; }
function setCfg_(k, v) { const s = JSON.stringify(v); if (!upd_('Config', 'key', k, { value: s })) add_('Config', { key: k, value: s }); }
function hash_(u, p) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + u + '|' + p)); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function days_(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 864e5); }
function fail_(m) { throw new Error(m); }
function mkSess_(o) { const t = Utilities.getUuid(); CacheService.getScriptCache().put(t, JSON.stringify(o), 21600); return t; }
function sess_(t) { const r = t && CacheService.getScriptCache().get(t); if (!r) fail_('หมดอายุการเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่'); return JSON.parse(r); }

/* ---------- API (เรียกจากหน้าเว็บ) ---------- */
function api(a, t, p) {
  const L = LockService.getScriptLock(); L.waitLock(25000); // ล็อกทีละคำสั่ง = กันจองซ้ำ/สมัครซ้ำ
  try { return { ok: true, d: route_(a, t, p || {}) }; }
  catch (e) { return { ok: false, m: String(e.message || e) }; }
  finally { L.releaseLock(); }
}

function route_(a, t, p) {
  if (a === 'register') return reg_(p);
  if (a === 'login') return login_(p);
  if (a === 'adminLogin') {
    if (String(p.u || '').trim().toUpperCase() !== ADMIN_U.toUpperCase() || String(p.pw || '').trim() !== ADMIN_P) fail_('Username หรือรหัสผ่านไม่ถูกต้อง');
    return { token: mkSess_({ u: ADMIN_U, admin: true }), user: { id: ADMIN_U, name: 'ผู้ดูแลระบบ' } };
  }
  if (a === 'pub') return { announce: cfg_().announce || null };
  const s = sess_(t);
  if (a === 'boot') return boot_(s);
  if (a === 'book') return book_(s, p);
  if (a === 'cancel') {
    const b = all_('Bookings').find(x => x.tx === p.tx) || fail_('ไม่พบรายการ');
    if (!s.admin && b.username !== s.u) fail_('ไม่มีสิทธิ์');
    if (!s.admin && b.ciStatus !== 'รอเข้าพัก') fail_('ยกเลิกไม่ได้ เนื่องจากเข้าพักแล้ว');
    upd_('Bookings', 'tx', p.tx, { status: 'cancelled', roomStatus: 'ยกเลิก' }); return 1;
  }
  if (!s.admin) fail_('ไม่มีสิทธิ์');
  if (a === 'bk') {
    const ok = ['ciStatus', 'coStatus', 'roomStatus', 'note', 'inspect'], q = {};
    ok.forEach(k => { if (p.patch[k] !== undefined) q[k] = p.patch[k]; });
    upd_('Bookings', 'tx', p.tx, q); return 1;
  }
  if (a === 'bkEdit') {
    const B = all_('Bookings');
    if (!B.some(x => x.tx === p.tx)) fail_('ไม่พบรายการ');
    const room = all_('Rooms').find(r => r.id === p.roomId) || fail_('ไม่พบห้อง');
    const n = days_(p.cin, p.cout);
    if (!(n >= 1)) fail_('วันที่ไม่ถูกต้อง (Check-out ต้องหลัง Check-in)');
    if (!String(p.name || '').trim()) fail_('กรุณากรอกชื่อ-นามสกุล');
    const st = p.roomStatus === 'ยกเลิก' ? 'cancelled' : 'active';
    if (st === 'active' && B.some(x => x.tx !== p.tx && x.roomId === room.id && x.status !== 'cancelled' && x.cin < p.cout && p.cin < x.cout)) fail_('ห้องนี้ถูกจองในช่วงวันที่เลือกแล้ว');
    upd_('Bookings', 'tx', p.tx, { roomId: room.id, roomName: room.name, name: String(p.name).trim(), phone: String(p.phone || '').trim(), shift: p.shift, channel: p.channel, cin: p.cin, cout: p.cout, nights: n, ciStatus: p.ciStatus, coStatus: p.coStatus, roomStatus: p.roomStatus, inspect: p.inspect, note: p.note || '', status: st });
    return 1;
  }
  if (a === 'insp') {
    add_('Inspections', { id: 'I' + Date.now(), date: p.date, roomId: p.roomId, inspector: p.inspector, result: p.result, tx: p.tx || '', note: p.note || '' });
    if (p.tx) upd_('Bookings', 'tx', p.tx, { inspect: p.result }); return 1;
  }
  if (a === 'room') {
    const o = { id: p.id, name: String(p.name || '').trim() || 'ห้อง', note: p.note || '', active: p.active === false || p.active === 'FALSE' ? 'FALSE' : 'TRUE' };
    if (p.id && all_('Rooms').some(r => r.id === p.id)) upd_('Rooms', 'id', p.id, o);
    else { o.id = 'R' + Date.now(); add_('Rooms', o); } return 1;
  }
  if (a === 'roomDel') {
    if (all_('Bookings').some(b => b.roomId === p.id)) fail_('ห้องนี้มีประวัติการจอง ไม่สามารถลบได้ ให้ปิดการใช้งานแทน');
    del_('Rooms', 'id', p.id); return 1;
  }
  if (a === 'cfg') { setCfg_('shifts', p.shifts); setCfg_('channels', p.channels); setCfg_('maxNights', Math.max(1, +p.maxNights || 7)); return 1; }
  if (a === 'ann') { if (JSON.stringify(p.ann || {}).length > 45000) fail_('ข้อความประกาศยาวเกินไป'); setCfg_('announce', p.ann || {}); return 1; }
  if (a === 'user') { upd_('Users', 'username', p.id, { status: p.status }); return 1; }
  if (a === 'userDel') { del_('Users', 'username', p.id); return 1; }
  fail_('คำสั่งไม่ถูกต้อง');
}

function reg_(p) {
  const u = String(p.u || '').trim().toLowerCase(), ph = String(p.ph || '').replace(/\D/g, '');
  if (!/^[a-z0-9_.-]{3,30}$/.test(u)) fail_('Username ต้องเป็นอังกฤษ/ตัวเลข 3–30 ตัว');
  if (String(p.pw || '').length < 6) fail_('รหัสผ่านอย่างน้อย 6 ตัว');
  if (!p.fn || !p.ln || !p.dp) fail_('กรุณากรอกข้อมูลให้ครบ');
  if (ph.length < 9 || ph.length > 10) fail_('เบอร์โทรศัพท์ไม่ถูกต้อง');
  const U = all_('Users');
  if (U.some(x => x.username.toLowerCase() === u)) fail_('Username นี้ถูกใช้แล้ว');
  if (U.some(x => x.phone === ph)) fail_('เบอร์โทรศัพท์นี้สมัครใช้งานแล้ว');
  add_('Users', { username: u, pass: hash_(u, p.pw), name: p.fn.trim() + ' ' + p.ln.trim(), dept: p.dp.trim(), phone: ph, status: 'active', createdAt: today_() });
  return 1;
}

function login_(p) {
  const u = String(p.u || '').trim().toLowerCase(), x = all_('Users').find(r => r.username.toLowerCase() === u);
  if (!x || x.pass !== hash_(x.username, p.pw)) fail_('Username หรือรหัสผ่านไม่ถูกต้อง');
  if (x.status === 'suspended') fail_('บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อแอดมิน');
  return { token: mkSess_({ u: x.username }), user: { id: x.username, name: x.name, dept: x.dept, phone: x.phone } };
}

function boot_(s) {
  const d = { rooms: all_('Rooms'), cfg: cfg_() }, B = all_('Bookings');
  if (s.admin) {
    d.bookings = B; d.insp = all_('Inspections');
    d.users = all_('Users').map(u => ({ id: u.username, name: u.name, dept: u.dept, phone: u.phone, status: u.status }));
  } else {
    const me = all_('Users').find(x => x.username === s.u);
    if (!me || me.status === 'suspended') fail_('หมดอายุการเข้าสู่ระบบ หรือบัญชีถูกระงับ');
    d.busy = B.filter(b => b.status !== 'cancelled').map(b => ({ roomId: b.roomId, cin: b.cin, cout: b.cout }));
    d.mine = B.filter(b => b.username === s.u);
  }
  return d;
}

function book_(s, p) {
  if (s.admin) fail_('กรุณาเข้าสู่ระบบด้วยบัญชีผู้ใช้เพื่อจอง');
  const me = all_('Users').find(x => x.username === s.u); if (!me || me.status === 'suspended') fail_('บัญชีถูกระงับ');
  const c = cfg_(), n = days_(p.cin, p.cout), room = all_('Rooms').find(r => r.id === p.roomId && r.active !== 'FALSE');
  if (!room) fail_('ไม่พบห้อง หรือห้องปิดใช้งาน');
  if (!(p.cin >= today_())) fail_('ไม่สามารถจองย้อนหลังได้');
  if (!(n >= 1)) fail_('วันที่ไม่ถูกต้อง');
  if (n > (c.maxNights || 7)) fail_('จองได้สูงสุด ' + (c.maxNights || 7) + ' คืนต่อครั้ง');
  const B = all_('Bookings');
  if (B.some(b => b.roomId === room.id && b.status !== 'cancelled' && b.cin < p.cout && p.cin < b.cout)) fail_('ห้องนี้ถูกจองในช่วงวันที่เลือกแล้ว กรุณาเลือกห้องอื่น');
  // รหัสรันต่อจากเลขสูงสุดที่มีอยู่ในชีต: ถ้าลบข้อมูลออก เลขจะย้อนกลับ (ลบหมด = เริ่ม BKH00001 ใหม่)
  const mx = B.reduce((m, b) => Math.max(m, parseInt(String(b.tx).replace(/\D/g, ''), 10) || 0), 0);
  const tx = 'BKH' + ('00000' + (mx + 1)).slice(-5);
  add_('Bookings', { tx: tx, createdAt: today_(), roomId: room.id, roomName: room.name, inspect: 'รอตรวจ', roomStatus: 'จองแล้ว', name: me.name, username: me.username, phone: me.phone, shift: p.shift, channel: (c.channels || DEF.channels)[0], cin: p.cin, cout: p.cout, nights: n, ciStatus: 'รอเข้าพัก', coStatus: 'รอออก', note: p.note || '', status: 'active' });
  return tx;
}
