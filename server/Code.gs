/**
 * VÔ ƯU TỰ – Máy chủ dữ liệu trên Google Sheets (Google Apps Script) · bản 5 (thêm lời chào có sẵn khi gặp nhau)
 * (bản 3: đếm người đang có mặt; bản 4: ghi giờ vào từng lượt, người dùng thấy nhau trong Tự, trang quản trị xem ai đang online)
 * Lưu: đăng ký dùng thử, góp ý, câu hỏi Phật pháp, Tường tâm nguyện, thông báo, báo lỗi, lượt vào Tự.
 * Có các lệnh quản trị (cần mật khẩu) cho trang quan-tri.html.
 *
 * Cập nhật từ bản 1 (xem server/HUONG-DAN.txt, mục "Cập nhật máy chủ"):
 *  1. Dán đè toàn bộ file này vào Apps Script, bấm Lưu.
 *  2. Chọn hàm setup → Chạy. Hàm chỉ thêm các trang tính mới, không xóa dữ liệu cũ.
 *  3. Chọn hàm xemMatKhauQuanTri → Chạy → mở "Nhật ký thực thi" để xem mật khẩu trang quản trị.
 *  4. Triển khai → Quản lý các bản triển khai → bút chì (Chỉnh sửa) → Phiên bản: "Phiên bản mới" → Triển khai.
 *     (Giữ nguyên bản triển khai cũ để đường dẫn /exec KHÔNG đổi.)
 */

const OWNER_EMAIL = '';          // Điền email của bạn để nhận thư báo khi có đăng ký, câu hỏi, tâm nguyện chờ duyệt, báo lỗi. Để trống nếu không cần.
const MAX_CALLS_PER_10_MIN = 120; // Chống spam: số lần gọi tối đa của một người trong 10 phút.

const SHEETS = {
  DangKy: ['Mã người dùng', 'Thời gian đăng ký', 'Tên gọi', 'Độ tuổi', 'Tỉnh/thành', 'Liên hệ', 'Muốn dùng để', 'Thiết bị', 'Đồng ý lưu thông tin', 'Cập nhật lần cuối'],
  GopY: ['Thời gian', 'Mã người dùng', 'Tên gọi', 'Điểm (1-5)', 'Điều thích', 'Cần cải thiện'],
  HoiDap: ['Mã câu hỏi', 'Thời gian', 'Mã người hỏi', 'Câu hỏi', 'Cho phép đăng', 'Lời đáp (bạn viết vào đây)', 'Đăng công khai', 'Ngày trả lời'],
  TamNguyen: ['Mã', 'Thời gian', 'Mã người gửi', 'Lời nguyện', 'Cầu cho', 'Trạng thái', 'Số lượt niệm'],
  ThongBao: ['Mã', 'Thời gian', 'Nội dung', 'Đang hiện'],
  BaoLoi: ['Mã', 'Thời gian', 'Mã người dùng', 'Mô tả', 'Thiết bị', 'Lỗi kỹ thuật', 'Phiên bản', 'Đã xử lý'],
  LuotVao: ['Ngày', 'Mã người dùng', 'Thiết bị', 'Phiên bản', 'Giờ vào', 'Thời điểm']
};
const ST_WAIT = 'Chờ duyệt', ST_OK = 'Đã duyệt', ST_HIDE = 'Ẩn';

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (name) {
    const head = SHEETS[name];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(head);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#f3e3c0').setWrap(true);
    }
  });
  const hd = ss.getSheetByName('HoiDap');
  hd.getRange('G2:G1000').insertCheckboxes();
  hd.setColumnWidth(4, 380); hd.setColumnWidth(6, 420);
  hd.getRange('D2:F1000').setWrap(true);
  const dk = ss.getSheetByName('DangKy');
  dk.setColumnWidth(1, 110); dk.setColumnWidth(7, 260);
  const tn = ss.getSheetByName('TamNguyen');
  tn.setColumnWidth(4, 420); tn.getRange('D2:D2000').setWrap(true);
  tn.getRange('F2:F2000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([ST_WAIT, ST_OK, ST_HIDE], true).build());
  ss.getSheetByName('ThongBao').setColumnWidth(3, 460);
  ss.getSheetByName('ThongBao').getRange('D2:D500').insertCheckboxes();
  ss.getSheetByName('BaoLoi').setColumnWidth(4, 380);
  ss.getSheetByName('BaoLoi').getRange('H2:H2000').insertCheckboxes();
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  adminKey_();
  Logger.log('Đã chuẩn bị xong. Chạy hàm xemMatKhauQuanTri để xem mật khẩu trang quản trị.');
}

/** Xem mật khẩu trang quản trị (mở Nhật ký thực thi sau khi chạy). */
function xemMatKhauQuanTri() { Logger.log('Mật khẩu trang quản trị: ' + adminKey_()); }
/** Quên mật khẩu: chạy hàm này để tạo mật khẩu ngẫu nhiên mới (mật khẩu cũ hết hiệu lực ngay), xem trong Nhật ký thực thi.
 *  Muốn đặt mật khẩu do mình chọn: đăng nhập trang quản trị, bấm "Đổi mật khẩu". */
function doiMatKhauQuanTri() {
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_KEY');
  Logger.log('Mật khẩu mới: ' + adminKey_());
}
function adminKey_() {
  const pr = PropertiesService.getScriptProperties();
  let k = pr.getProperty('ADMIN_KEY');
  if (!k) {
    const a = 'abcdefghjkmnpqrstuvwxyz23456789';
    k = 'vut-';
    for (let i = 0; i < 12; i++) k += a.charAt(Math.floor(Math.random() * a.length));
    pr.setProperty('ADMIN_KEY', k);
  }
  return k;
}

function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return handle_(p);
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0009\u000b-\u001f\u007f]/g, ' ').trim().slice(0, max || 500); }
function cell_(v, max) { v = clean_(v, max || 1500); return /^[=+\-@]/.test(v) ? "'" + v : v; } // chống chèn công thức vào Sheet
function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh && SHEETS[name]) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); }
  return sh;
}
function newId_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 7); } // có chữ cái đầu để Sheets không tự đổi thành số
function ts_(v) { const t = new Date(v).getTime(); return isNaN(t) ? 0 : t; }
function day_(d) { return Utilities.formatDate(d || new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd'); }

function rate_(cid) {
  const c = CacheService.getScriptCache(), k = 'r_' + cid;
  const n = Number(c.get(k) || 0) + 1;
  c.put(k, String(n), 600);
  return n <= MAX_CALLS_PER_10_MIN;
}

function handle_(p) {
  try {
    const action = clean_(p.action, 20);
    if (action === 'faq') return out_({ ok: true, items: faq_() });
    if (action.indexOf('a_') === 0) return out_(admin_(action, p));
    const cid = clean_(p.cid, 64);
    if (!/^[a-z0-9-]{8,64}$/i.test(cid)) return out_({ ok: false, error: 'bad_cid' });
    if (!rate_(cid)) return out_({ ok: false, error: 'rate' });
    if (action === 'getReg') return out_({ ok: true, reg: getReg_(cid) });
    if (action === 'myq') return out_({ ok: true, items: myq_(cid) });
    if (action === 'wall') return out_({ ok: true, items: wallList_(), mine: wallMine_(cid) });
    if (action === 'hello') return out_(hello_(cid, p));
    if (action === 'here') return out_(here_(cid, p));
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (action === 'register') return out_(register_(cid, p));
      if (action === 'feedback') return out_(feedback_(cid, p));
      if (action === 'ask') return out_(ask_(cid, p));
      if (action === 'wallPost') return out_(wallPost_(cid, p));
      if (action === 'wallNiem') return out_(wallNiem_(cid, p));
      if (action === 'bug') return out_(bug_(cid, p));
    } finally { lock.releaseLock(); }
    return out_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out_({ ok: false, error: 'server', detail: String(err) });
  }
}

function findRow_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function rows_(name) {
  const sh = sheet_(name), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, SHEETS[name].length).getValues();
}

/* ---------- Người dùng ---------- */

function register_(cid, p) {
  const name = clean_(p.name, 60);
  if (!name || p.consent !== true) return { ok: false, error: 'invalid' };
  const sh = sheet_('DangKy');
  const uses = Array.isArray(p.uses) ? p.uses.map(function (u) { return clean_(u, 60); }).join('; ') : '';
  const now = new Date();
  const row = findRow_(sh, cid);
  const vals = [cid, now, cell_(name, 60), cell_(p.age, 20), cell_(p.prov, 60), cell_(p.contact, 100), cell_(uses, 400), cell_(p.dev, 30), 'Có', now];
  if (row > 0) {
    vals[1] = sh.getRange(row, 2).getValue();
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  } else {
    sh.appendRow(vals);
    notify_('Vô Ưu Tự: có người đăng ký dùng thử', name + ' (' + clean_(p.prov, 60) + ') vừa đăng ký dùng thử.\nLiên hệ: ' + clean_(p.contact, 100));
  }
  return { ok: true };
}

function getReg_(cid) {
  const sh = sheet_('DangKy'), row = findRow_(sh, cid);
  if (row < 0) return null;
  const v = sh.getRange(row, 1, 1, 10).getValues()[0];
  return { name: v[2], age: v[3], prov: v[4], contact: v[5], uses: String(v[6] || '').split('; ').filter(String), dev: v[7], consent: v[8] === 'Có', at: ts_(v[1]) };
}

function feedback_(cid, p) {
  const reg = getReg_(cid);
  const star = Math.max(0, Math.min(5, Number(p.star) || 0));
  sheet_('GopY').appendRow([new Date(), cid, cell_(reg ? reg.name : '', 60), star || '', cell_(p.like, 800), cell_(p.fix, 800)]);
  return { ok: true };
}

function ask_(cid, p) {
  const q = clean_(p.q, 1200);
  if (q.length < 3) return { ok: false, error: 'invalid' };
  const id = newId_('q');
  const sh = sheet_('HoiDap');
  sh.appendRow([id, new Date(), cid, cell_(q, 1200), p.share === true ? 'Có' : 'Không', '', false, '']);
  sh.getRange(sh.getLastRow(), 7).insertCheckboxes();
  notify_('Vô Ưu Tự: có câu hỏi mới', q + '\n\nVào trang quản trị hoặc Google Sheet (trang "HoiDap") để trả lời.');
  return { ok: true, id: id };
}

function myq_(cid) {
  return rows_('HoiDap').filter(function (r) { return r[2] === cid; }).slice(-20).map(function (r) {
    return { id: r[0], at: ts_(r[1]), q: String(r[3]), a: String(r[5] || '') };
  });
}

function faq_() {
  return rows_('HoiDap').filter(function (r) { return r[4] === 'Có' && r[6] === true && String(r[5] || '').trim(); })
    .slice(-30).reverse().map(function (r) { return { q: String(r[3]), a: String(r[5]) }; });
}

/** Lời chào khi vào Tự: ghi mỗi lượt vào (cách nhau từ 20 phút trở lên tính là lượt mới) kèm giờ vào, trả về thông báo đang hiện. */
function hello_(cid, p) {
  const c = CacheService.getScriptCache(), k = 'v_' + cid;
  if (!c.get(k)) {
    c.put(k, '1', 1200);
    const sh = sheet_('LuotVao');
    if (!c.get('lv_head') && sh.getLastColumn() < SHEETS.LuotVao.length) sh.getRange(1, 1, 1, SHEETS.LuotVao.length).setValues([SHEETS.LuotVao]).setFontWeight('bold').setBackground('#f3e3c0');
    c.put('lv_head', '1', 21600);
    const now = new Date();
    sh.appendRow([day_(now), cid, cell_(p.dev, 30), cell_(p.ver, 30), "'" + Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'HH:mm'), now]);
  }
  return { ok: true, ann: annList_(true) };
}

/** Hiện diện thật: mỗi người báo "đang ở trong Tự" khoảng 15 giây một lần, kèm vị trí, dáng, trang phục (không có thông tin cá nhân).
 *  Trả về số người có mặt và vị trí những người khác để mọi người thấy nhau. Ai im lặng quá 70 giây coi như đã rời đi. */
const PRES_TTL = 70000;
function presRead_(c) { try { return JSON.parse(c.get('pres') || '{}'); } catch (err) { return {}; } }
function here_(cid, p) {
  const c = CacheService.getScriptCache(), lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const m = presRead_(c), now = Date.now();
    Object.keys(m).forEach(function (k) { if (!Array.isArray(m[k]) || now - m[k][0] > PRES_TTL) delete m[k]; });
    const q = p.s || {}, num = function (v, lo, hi) { v = Number(v); return isFinite(v) ? Math.max(lo, Math.min(hi, Math.round(v * 100) / 100)) : 0; };
    const st = {
      i: String(q.i || '').replace(/[^a-z0-9]/gi, '').slice(0, 10),
      x: num(q.x, -60, 60), z: num(q.z, -60, 60), y: num(q.y, -20, 20),
      o: ['lam', 'aodai', 'tang'].indexOf(q.o) >= 0 ? q.o : 'lam', f: q.f ? 1 : 0,
      p: ['stand', 'sit', 'lay', 'chao'].indexOf(q.p) >= 0 ? q.p : 'stand',
      n: clean_(q.n, 20).replace(/[<>"&]/g, ''),
      m: [0, 1, 2].indexOf(Number(q.m)) >= 0 ? Number(q.m) : 0
    };
    const prev = m[cid];
    m[cid] = [now, p.chant === true ? 1 : 0, prev ? prev[2] : now, st];
    const s = JSON.stringify(m);
    if (s.length < 95000) c.put('pres', s, 600);
    let n = 0, ch = 0;
    const peers = [];
    Object.keys(m).forEach(function (k) {
      n++; if (m[k][1]) ch++;
      if (k !== cid && m[k][3] && m[k][3].i && peers.length < 40) { const o = m[k][3]; peers.push({ i: o.i, x: o.x, z: o.z, y: o.y, o: o.o, f: o.f, p: o.p, n: o.n, m: o.m || 0, a: now - m[k][0] }); }
    });
    return { ok: true, n: n, chant: ch, peers: peers };
  } finally { lock.releaseLock(); }
}
function zone_(x, z) {
  const d = function (a, b) { return Math.sqrt((x - a) * (x - a) + (z - b) * (z - b)); };
  if (z < -8 && Math.abs(x) < 11) return 'Chánh điện';
  if (d(-8, 8) < 5.5) return 'Hồ sen';
  if (d(13, 0.4) < 4) return 'Cội Bồ đề';
  if (d(8.6, 12.5) < 3.2) return 'Tượng Quán Âm';
  if (d(-8.6, -2.5) < 3.2) return 'Gác chuông';
  if (d(6.6, 10.2) < 2) return 'Bàn xin xăm';
  if (d(5.5, 4) < 2.5) return 'Đỉnh hương';
  if (z > 19) return 'Cổng Tam Quan';
  return 'Sân Tự';
}

function annList_(onlyOn) {
  return rows_('ThongBao').filter(function (r) { return r[0] && (!onlyOn || r[3] === true); })
    .map(function (r) { return { id: String(r[0]), at: ts_(r[1]), text: String(r[2]), on: r[3] === true }; }).reverse();
}

/* ---------- Tường tâm nguyện ---------- */

function wallList_() {
  return rows_('TamNguyen').filter(function (r) { return r[5] === ST_OK; }).slice(-60).reverse()
    .map(function (r) { return { id: String(r[0]), at: ts_(r[1]), t: String(r[3]), w: String(r[4]), n: Number(r[6]) || 0 }; });
}
function wallMine_(cid) {
  return rows_('TamNguyen').filter(function (r) { return r[2] === cid && r[5] !== ST_HIDE; }).slice(-10).reverse()
    .map(function (r) { return { id: String(r[0]), at: ts_(r[1]), t: String(r[3]), w: String(r[4]), st: String(r[5]), n: Number(r[6]) || 0 }; });
}
function wallPost_(cid, p) {
  const t = clean_(p.t, 300);
  if (t.length < 2) return { ok: false, error: 'invalid' };
  const mine = rows_('TamNguyen').filter(function (r) { return r[2] === cid && ts_(r[1]) > Date.now() - 86400000; });
  if (mine.length >= 5) return { ok: false, error: 'rate' };
  const id = newId_('t');
  sheet_('TamNguyen').appendRow([id, new Date(), cid, cell_(t, 300), cell_(p.w, 40), ST_WAIT, 0]);
  notify_('Vô Ưu Tự: có tâm nguyện chờ duyệt', t + '\n\nVào trang quản trị để duyệt.');
  return { ok: true, id: id };
}
function wallNiem_(cid, p) {
  const id = clean_(p.id, 20), c = CacheService.getScriptCache(), k = 'n_' + cid + '_' + id;
  if (c.get(k)) return { ok: true, dup: true };
  const sh = sheet_('TamNguyen'), row = findRow_(sh, id);
  if (row < 0 || sh.getRange(row, 6).getValue() !== ST_OK) return { ok: false, error: 'not_found' };
  const cell = sh.getRange(row, 7), n = (Number(cell.getValue()) || 0) + 1;
  cell.setValue(n);
  c.put(k, '1', 21600);
  return { ok: true, n: n };
}

/* ---------- Báo lỗi ---------- */

function bug_(cid, p) {
  const t = clean_(p.text, 1200);
  if (t.length < 3) return { ok: false, error: 'invalid' };
  const id = newId_('b');
  sheet_('BaoLoi').appendRow([id, new Date(), cid, cell_(t, 1200), cell_(p.dev, 200), cell_(p.err, 1500), cell_(p.ver, 30), false]);
  notify_('Vô Ưu Tự: có báo lỗi mới', t + '\n\nThiết bị: ' + clean_(p.dev, 200));
  return { ok: true };
}

/* ---------- Quản trị (trang quan-tri.html) ---------- */

function admin_(action, p) {
  const c = CacheService.getScriptCache();
  const fails = Number(c.get('afail') || 0);
  if (fails >= 10) return { ok: false, error: 'locked' };
  if (clean_(p.key, 60) !== adminKey_()) { c.put('afail', String(fails + 1), 900); return { ok: false, error: 'bad_key' }; }
  if (action === 'a_login') return { ok: true };
  if (action === 'a_setKey') {
    const nk = String(p.newKey == null ? '' : p.newKey).trim();
    if (nk.length < 8 || nk.length > 60 || /\s/.test(nk)) return { ok: false, error: 'weak_key' };
    PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', nk);
    return { ok: true };
  }
  if (action === 'a_stats') return { ok: true, stats: stats_() };
  if (action === 'a_qs') return { ok: true, items: rows_('HoiDap').slice(-300).reverse().map(function (r) {
    return { id: String(r[0]), at: ts_(r[1]), q: String(r[3]), share: r[4] === 'Có', a: String(r[5] || ''), pub: r[6] === true, ad: ts_(r[7]) }; }) };
  if (action === 'a_wall') return { ok: true, items: rows_('TamNguyen').slice(-400).reverse().map(function (r) {
    return { id: String(r[0]), at: ts_(r[1]), t: String(r[3]), w: String(r[4]), st: String(r[5]), n: Number(r[6]) || 0 }; }) };
  if (action === 'a_live') {
    const m = presRead_(CacheService.getScriptCache()), now = Date.now(), regs = {}, devs = {};
    rows_('DangKy').forEach(function (r) { regs[r[0]] = String(r[2]); });
    rows_('LuotVao').forEach(function (r) { devs[r[1]] = String(r[2]); });
    const items = Object.keys(m).filter(function (k) { return Array.isArray(m[k]) && now - m[k][0] <= PRES_TTL; }).map(function (k) {
      const v = m[k], st = v[3] || {};
      return { id: k.slice(0, 8), name: regs[k] || '', shown: st.n || '', dev: devs[k] || '', since: v[2] || v[0], last: v[0], chant: !!v[1], zone: st.i ? zone_(st.x, st.z) : '', pose: st.p || '', outfit: st.o || '' };
    }).sort(function (a, b) { return a.since - b.since; });
    return { ok: true, items: items };
  }
  if (action === 'a_visits') {
    const regs = {};
    rows_('DangKy').forEach(function (r) { regs[r[0]] = String(r[2]); });
    const all = rows_('LuotVao'), today = day_(), uniq = {}, uToday = {};
    let sToday = 0;
    all.forEach(function (r) { const d = r[0] instanceof Date ? day_(r[0]) : String(r[0]); uniq[r[1]] = 1; if (d === today) { sToday++; uToday[r[1]] = 1; } });
    const items = all.slice(-300).reverse().map(function (r) {
      const d = r[0] instanceof Date ? day_(r[0]) : String(r[0]);
      return { day: d, time: String(r[4] || '').replace(/^'/, ''), at: ts_(r[5]), id: String(r[1]).slice(0, 8), name: regs[r[1]] || '', dev: String(r[2]), ver: String(r[3]) };
    });
    return { ok: true, items: items, sum: { sessionsToday: sToday, peopleToday: Object.keys(uToday).length, peopleAll: Object.keys(uniq).length, sessionsAll: all.length } };
  }
  if (action === 'a_ann') return { ok: true, items: annList_(false) };
  if (action === 'a_fb') return { ok: true,
    fb: rows_('GopY').slice(-300).reverse().map(function (r) { return { at: ts_(r[0]), name: String(r[2]), star: Number(r[3]) || 0, like: String(r[4]), fix: String(r[5]) }; }),
    bugs: rows_('BaoLoi').slice(-300).reverse().map(function (r) { return { id: String(r[0]), at: ts_(r[1]), cid: String(r[2]), text: String(r[3]), dev: String(r[4]), err: String(r[5]), ver: String(r[6]), done: r[7] === true }; }) };
  if (action === 'a_reg') return { ok: true, items: rows_('DangKy').reverse().map(function (r) {
    return { at: ts_(r[1]), name: String(r[2]), age: String(r[3]), prov: String(r[4]), contact: String(r[5]), uses: String(r[6]), dev: String(r[7]) }; }) };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (action === 'a_answer') {
      const sh = sheet_('HoiDap'), row = findRow_(sh, clean_(p.id, 20));
      if (row < 0) return { ok: false, error: 'not_found' };
      const a = clean_(p.a, 5000), share = sh.getRange(row, 5).getValue() === 'Có';
      sh.getRange(row, 6, 1, 3).setValues([[cell_(a, 5000), !!(p.pub && share && a), a ? new Date() : '']]);
      return { ok: true };
    }
    if (action === 'a_wallSet') {
      const sh = sheet_('TamNguyen'), row = findRow_(sh, clean_(p.id, 20));
      if (row < 0) return { ok: false, error: 'not_found' };
      const st = [ST_WAIT, ST_OK, ST_HIDE].indexOf(p.st) >= 0 ? p.st : ST_WAIT;
      if (p.t != null && clean_(p.t, 300).length >= 2) sh.getRange(row, 4).setValue(cell_(p.t, 300));
      sh.getRange(row, 6).setValue(st);
      return { ok: true };
    }
    if (action === 'a_annAdd') {
      const t = clean_(p.text, 600);
      if (t.length < 2) return { ok: false, error: 'invalid' };
      const sh = sheet_('ThongBao');
      sh.appendRow([newId_('a'), new Date(), cell_(t, 600), true]);
      sh.getRange(sh.getLastRow(), 4).insertCheckboxes().setValue(true);
      return { ok: true };
    }
    if (action === 'a_annSet') {
      const sh = sheet_('ThongBao'), row = findRow_(sh, clean_(p.id, 20));
      if (row < 0) return { ok: false, error: 'not_found' };
      sh.getRange(row, 4).setValue(p.on === true);
      return { ok: true };
    }
    if (action === 'a_purge') {
      const cid = clean_(p.cid, 64);
      if (!/^[a-z0-9-]{8,64}$/i.test(cid)) return { ok: false, error: 'invalid' };
      const cols = { DangKy: 1, GopY: 2, HoiDap: 3, TamNguyen: 3, LuotVao: 2, BaoLoi: 3 };
      let n = 0;
      Object.keys(cols).forEach(function (name) {
        const sh = sheet_(name), last = sh.getLastRow();
        if (last < 2) return;
        const ids = sh.getRange(2, cols[name], last - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === cid) { sh.deleteRow(i + 2); n++; }
      });
      return { ok: true, n: n };
    }
    if (action === 'a_bugSet') {
      const sh = sheet_('BaoLoi'), row = findRow_(sh, clean_(p.id, 20));
      if (row < 0) return { ok: false, error: 'not_found' };
      sh.getRange(row, 8).setValue(p.done === true);
      return { ok: true };
    }
  } finally { lock.releaseLock(); }
  return { ok: false, error: 'unknown_action' };
}

function stats_() {
  const visits = rows_('LuotVao'), byDay = {};
  visits.forEach(function (r) {
    const d = r[0] instanceof Date ? day_(r[0]) : String(r[0]);
    (byDay[d] = byDay[d] || {})[r[1]] = 1;
  });
  const days = [];
  for (let i = 13; i >= 0; i--) {
    const d = day_(new Date(Date.now() - i * 86400000));
    days.push({ d: d, n: Object.keys(byDay[d] || {}).length });
  }
  const all = {};
  visits.forEach(function (r) { all[r[1]] = 1; });
  const fb = rows_('GopY'), stars = fb.map(function (r) { return Number(r[3]) || 0; }).filter(Boolean);
  const qs = rows_('HoiDap'), wall = rows_('TamNguyen'), bugs = rows_('BaoLoi');
  return {
    reg: rows_('DangKy').length,
    fb: fb.length,
    star: stars.length ? Math.round(stars.reduce(function (a, b) { return a + b; }, 0) / stars.length * 10) / 10 : 0,
    qTotal: qs.length,
    qWait: qs.filter(function (r) { return !String(r[5] || '').trim(); }).length,
    wallWait: wall.filter(function (r) { return r[5] === ST_WAIT; }).length,
    wallOk: wall.filter(function (r) { return r[5] === ST_OK; }).length,
    niem: wall.reduce(function (a, r) { return a + (Number(r[6]) || 0); }, 0),
    bugOpen: bugs.filter(function (r) { return r[7] !== true; }).length,
    people: Object.keys(all).length,
    days: days
  };
}

/** Tự ghi ngày trả lời khi bạn viết lời đáp trực tiếp trong Sheet (trình kích hoạt onEdit đơn giản). */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== 'HoiDap' || e.range.getColumn() !== 6 || e.range.getRow() < 2) return;
    sh.getRange(e.range.getRow(), 8).setValue(e.value ? new Date() : '');
  } catch (err) {}
}

function notify_(subject, body) {
  if (!OWNER_EMAIL) return;
  try { MailApp.sendEmail(OWNER_EMAIL, subject, body); } catch (err) {}
}
