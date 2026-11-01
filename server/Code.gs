/**
 * VÔ ƯU TỰ – Máy chủ dữ liệu trên Google Sheets (Google Apps Script)
 * Lưu: đăng ký dùng thử, góp ý, câu hỏi Phật pháp; trả về lời đáp và mục Hỏi đáp chung.
 *
 * Cách dùng (chi tiết trong HUONG-DAN):
 *  1. Tạo một Google Sheet mới → Tiện ích mở rộng → Apps Script → dán toàn bộ file này.
 *  2. Chọn hàm setup → Chạy (cấp quyền lần đầu).
 *  3. Triển khai → Tùy chọn triển khai mới → Ứng dụng web; Thực thi với tư cách: Tôi; Người có quyền truy cập: Bất kỳ ai.
 *  4. Sao chép URL kết thúc bằng /exec, dán vào file config.js của website.
 *
 * Trả lời câu hỏi: mở trang tính "HoiDap", viết vào cột "Lời đáp", tick cột "Đăng công khai" nếu muốn đưa vào Hỏi đáp chung
 * (chỉ đăng được khi người hỏi cho phép).
 */

const OWNER_EMAIL = '';          // Điền email của bạn để nhận thư báo khi có đăng ký hoặc câu hỏi mới. Để trống nếu không cần.
const MAX_CALLS_PER_10_MIN = 40; // Chống spam: số lần gửi tối đa của một người trong 10 phút.

const SHEETS = {
  DangKy: ['Mã người dùng', 'Thời gian đăng ký', 'Tên gọi', 'Độ tuổi', 'Tỉnh/thành', 'Liên hệ', 'Muốn dùng để', 'Thiết bị', 'Đồng ý lưu thông tin', 'Cập nhật lần cuối'],
  GopY: ['Thời gian', 'Mã người dùng', 'Tên gọi', 'Điểm (1-5)', 'Điều thích', 'Cần cải thiện'],
  HoiDap: ['Mã câu hỏi', 'Thời gian', 'Mã người hỏi', 'Câu hỏi', 'Cho phép đăng', 'Lời đáp (bạn viết vào đây)', 'Đăng công khai', 'Ngày trả lời']
};

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
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return handle_(p);
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max || 500); }
function cell_(v, max) { v = clean_(v, max || 1500); return /^[=+\-@]/.test(v) ? "'" + v : v; } // chống chèn công thức vào Sheet
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

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
    const cid = clean_(p.cid, 64);
    if (!/^[a-z0-9-]{8,64}$/i.test(cid)) return out_({ ok: false, error: 'bad_cid' });
    if (!rate_(cid)) return out_({ ok: false, error: 'rate' });
    if (action === 'getReg') return out_({ ok: true, reg: getReg_(cid) });
    if (action === 'myq') return out_({ ok: true, items: myq_(cid) });
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (action === 'register') return out_(register_(cid, p));
      if (action === 'feedback') return out_(feedback_(cid, p));
      if (action === 'ask') return out_(ask_(cid, p));
    } finally { lock.releaseLock(); }
    return out_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out_({ ok: false, error: 'server', detail: String(err) });
  }
}

function findRow_(sh, cid) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === cid) return i + 2;
  return -1;
}

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
  return { name: v[2], age: v[3], prov: v[4], contact: v[5], uses: String(v[6] || '').split('; ').filter(String), dev: v[7], consent: v[8] === 'Có', at: new Date(v[1]).getTime() };
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
  const id = Utilities.getUuid().slice(0, 8);
  const sh = sheet_('HoiDap');
  sh.appendRow([id, new Date(), cid, cell_(q, 1200), p.share === true ? 'Có' : 'Không', '', false, '']);
  sh.getRange(sh.getLastRow(), 7).insertCheckboxes();
  notify_('Vô Ưu Tự: có câu hỏi mới', q + '\n\nMở Google Sheet, trang "HoiDap" để trả lời.');
  return { ok: true, id: id };
}

function rows_(name) {
  const sh = sheet_(name), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}

function myq_(cid) {
  return rows_('HoiDap').filter(function (r) { return r[2] === cid; }).slice(-20).map(function (r) {
    return { id: r[0], at: new Date(r[1]).getTime(), q: String(r[3]), a: String(r[5] || '') };
  });
}

function faq_() {
  return rows_('HoiDap').filter(function (r) { return r[4] === 'Có' && r[6] === true && String(r[5] || '').trim(); })
    .slice(-30).reverse().map(function (r) { return { q: String(r[3]), a: String(r[5]) }; });
}

/** Tự ghi ngày trả lời khi bạn viết lời đáp trong Sheet (trình kích hoạt onEdit đơn giản). */
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
