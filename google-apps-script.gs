/**
 * Fohepta Plus – nhận đơn vào Google Sheet + trả dữ liệu THẬT cho landing page
 * (popup "vừa có người mua", số đơn hôm nay, số người đang xem).
 *
 * CÁCH CÀI (5 phút):
 * 1. Tạo một Google Sheet mới → menu Extensions → Apps Script.
 * 2. Xoá code mẫu, dán toàn bộ file này vào → Save.
 * 3. Deploy → New deployment → Select type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    → Deploy → cấp quyền → copy "Web app URL".
 * 4. Trong index.html, dán URL đó vào cả ORDER_ENDPOINT và SOCIAL_FEED_URL.
 * 5. Mỗi lần sửa script: Deploy → Manage deployments → Edit → Version: New version → Deploy.
 *
 * Dữ liệu công khai trả cho trang CHỈ gồm: tên + chữ cái đầu của họ, thành phố, gói, thời gian.
 * Không bao giờ trả số điện thoại hay địa chỉ. Nên ghi rõ việc hiển thị này trong chính sách bảo mật.
 */

const SHEET_NAME = 'Orders';
const TZ = 'Asia/Manila';
const HEADERS = ['Thời gian', 'Họ tên', 'SĐT', 'City/Province', 'Địa chỉ', 'Gói', 'Số lon', 'Tổng (PHP)',
  'Thanh toán', 'Event ID', 'utm_source', 'utm_campaign', 'utm_content', 'fbclid', 'Trang'];

/** Nhận đơn hàng và ping "đang xem" từ landing page (POST form-urlencoded). */
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') {
    ping_(p.sid);
    return json_({ ok: true });
  }
  if (!p.name || !p.phone) return json_({ ok: false, error: 'missing fields' });
  sheet_().appendRow([
    new Date(), p.name, "'" + p.phone, p.city || '', p.address || '', p.package || '',
    Number(p.cans) || '', Number(p.total) || '', p.payment || 'COD', p.event_id || '',
    p.utm_source || '', p.utm_campaign || '', p.utm_content || '', p.fbclid || '', p.page || ''
  ]);
  return json_({ ok: true });
}

/** Trả dữ liệu thật cho trang: { recent: [...], ordersToday: n, viewers: n } */
function doGet() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const out = { recent: [], ordersToday: 0, viewers: viewers_() };
  if (last > 1) {
    const n = Math.min(1000, last - 1);
    const rows = sh.getRange(last - n + 1, 1, n, 6).getValues();
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    rows.forEach(function (r) {
      if (r[0] instanceof Date && Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') === today) out.ordersToday++;
    });
    out.recent = rows.slice(-15).reverse()
      .filter(function (r) { return r[0] instanceof Date && r[1]; })
      .map(function (r) {
        return {
          name: shortName_(r[1]),
          city: String(r[3] || '').split(',')[0].trim(),
          package: String(r[5] || ''),
          ts: r[0].toISOString()
        };
      });
  }
  return json_(out);
}

/* ---------- "Đang xem": đếm số phiên khác nhau đã ping trong 3 phút gần nhất ---------- */
function ping_(sid) {
  if (!sid) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const cache = CacheService.getScriptCache();
    const key = 'v' + Math.floor(Date.now() / 60000);
    const ids = JSON.parse(cache.get(key) || '[]');
    if (ids.indexOf(sid) < 0 && ids.length < 4000) {
      ids.push(sid);
      cache.put(key, JSON.stringify(ids), 300);
    }
  } finally {
    lock.releaseLock();
  }
}

function viewers_() {
  const cache = CacheService.getScriptCache();
  const m = Math.floor(Date.now() / 60000);
  const keys = ['v' + m, 'v' + (m - 1), 'v' + (m - 2)];
  const got = cache.getAll(keys);
  const seen = {};
  keys.forEach(function (k) { JSON.parse(got[k] || '[]').forEach(function (s) { seen[s] = 1; }); });
  return Object.keys(seen).length;
}

/* ---------- Helpers ---------- */
function shortName_(full) {
  const parts = String(full).trim().split(/\s+/);
  const cap = function (w) { return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); };
  return parts.length > 1 ? cap(parts[0]) + ' ' + parts[parts.length - 1].charAt(0).toUpperCase() + '.' : cap(parts[0]);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
