/**
 * GadgetVilla Dashboard - Google Apps Script Web App (Data Layer)
 * doGet  : READ  - aggregates Data_ALL + Detail_ALL into the dashboard's schema.
 * doPost : WRITE - Admin CRUD (create / update / delete).
 *
 * Standalone-safe: opens the sheet by ID (Script Property SHEET_ID overrides).
 * Optional monthly target: Script Property SALES_TARGET (number) -> fills goal/percent.
 */
var SHEET_ID = '15u6K6fl4VukQc8SdKUvnlNcq5PhrLqPzp78L3hc3Cgw';

function _ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || SHEET_ID;
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function _prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function _token_() { return _prop_('API_TOKEN'); }
function _checkToken_(p) { var e = _token_(); if (!e) return true; return p === e; }
function _json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- small utils ---------- */
function _num_(x) {
  if (x === '' || x === null || x === undefined) return 0;
  if (typeof x === 'number') return x;
  var n = parseFloat(String(x).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}
function _yes_(x) {
  if (x === true) return true;
  var s = String(x).trim().toLowerCase();
  return s === 'ใช่' || s === 'y' || s === 'yes' || s === 'true' || s === '1' || s.indexOf('ชำระ') >= 0;
}
function _fmt_(n) {
  n = Math.round(Number(n) || 0);
  var s = String(Math.abs(n)), out = '';
  while (s.length > 3) { out = ',' + s.slice(-3) + out; s = s.slice(0, -3); }
  return (n < 0 ? '-' : '') + s + out;
}
function _pct_(x) { return (Math.round((Number(x) || 0) * 100) / 100).toFixed(2) + '%'; }
var _THMON = {'01':'ม.ค.','02':'ก.พ.','03':'มี.ค.','04':'เม.ย.','05':'พ.ค.','06':'มิ.ย.',
              '07':'ก.ค.','08':'ส.ค.','09':'ก.ย.','10':'ต.ค.','11':'พ.ย.','12':'ธ.ค.'};
function _thMon_(ym) { var mm = String(ym).split('-')[1]; return _THMON[mm] || String(ym); }

/* read only the named columns of a sheet (avoids loading huge unused columns) */
function _cols_(sheet, names) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var idx = {}; header.forEach(function (h, i) { idx[h] = i; });
  var n = lastRow - 1, col = {};
  names.forEach(function (nm) {
    var c = idx[nm];
    if (c === undefined || n < 1) { col[nm] = []; return; }
    col[nm] = sheet.getRange(2, c + 1, n, 1).getValues().map(function (r) { return r[0]; });
  });
  return { n: (n < 1 ? 0 : n), col: col };
}

/* ---------- column names ---------- */
var C = {
  MONTH: 'เดือน', AGENT: 'ชื่อบริการลูกค้า',
  CHATS: 'จำนวนสนทนาการเข้าถึงทั้งหมด',
  RECEIVED: 'จำนวนสนทนาที่ได้ต้อนรับจากฝ่ายบริการลูกค้า',
  GUIDED: 'จำนวนผู้ซื้อที่ชี้นำสั่งซื้อ',
  PAID_ORD: 'จำนวนคำสั่งซื้อที่ชี้นำชำระ',
  PAID_AMT: 'ยอดชำระคำสั่งซื้อที่ชี้นำชำระ',
  PAYCONV: 'อัตราการแปลงการชำระเงิน',
  W1: 'จำนวนตอบสนองภายใน 1 นาที', W5: 'จำนวนตอบสนองภายใน 5 นาที',
  W10: 'จำนวนตอบสนองภายใน 10 นาที', W30: 'จำนวนตอบสนองภายใน 30 นาที',
  NEG: 'จำนวนรีวิวเชิงลบระดับกลางสำหรับร้านค้าที่มีสิทธิ์ของฝ่ายบริการลูกค้า',
  D_MONTH: 'เดือน', D_SHOP: 'ชื่อร้าน', D_PAID: 'หลังต้อนรับได้ชำระหรือไม่'
};

function _buildDashboard_(ss) {
  var target = _num_(_prop_('SALES_TARGET'));
  var out = {};

  /* ===== Data_ALL (per agent per month) ===== */
  var dS = ss.getSheetByName('Data_ALL');
  if (dS) {
    var D = _cols_(dS, [C.MONTH, C.AGENT, C.CHATS, C.RECEIVED, C.GUIDED, C.PAID_ORD,
                        C.PAID_AMT, C.PAYCONV, C.W1, C.W5, C.W10, C.W30, C.NEG]);
    var n = D.n, m = D.col;
    var months = [];
    for (var i = 0; i < n; i++) { var mo = m[C.MONTH][i]; if (mo && months.indexOf(mo) < 0) months.push(mo); }
    months.sort();
    var latest = months[months.length - 1];

    // latest-month totals
    var t = { sales: 0, orders: 0, chats: 0, recv: 0, guided: 0, neg: 0, w1: 0, w5: 0, w10: 0, w30: 0 };
    var staff = {};
    var trendByMonth = {}; months.forEach(function (mm) { trendByMonth[mm] = 0; });
    for (var r = 0; r < n; r++) {
      var mo2 = m[C.MONTH][r];
      trendByMonth[mo2] = (trendByMonth[mo2] || 0) + _num_(m[C.PAID_AMT][r]);
      if (mo2 !== latest) continue;
      t.sales += _num_(m[C.PAID_AMT][r]); t.orders += _num_(m[C.PAID_ORD][r]);
      t.chats += _num_(m[C.CHATS][r]);   t.recv += _num_(m[C.RECEIVED][r]);
      t.guided += _num_(m[C.GUIDED][r]);  t.neg += _num_(m[C.NEG][r]);
      t.w1 += _num_(m[C.W1][r]); t.w5 += _num_(m[C.W5][r]); t.w10 += _num_(m[C.W10][r]); t.w30 += _num_(m[C.W30][r]);
      var a = m[C.AGENT][r] || '-';
      if (!staff[a]) staff[a] = { name: a, chats: 0, sales: 0, conv: 0 };
      staff[a].chats += _num_(m[C.CHATS][r]);
      staff[a].sales += _num_(m[C.PAID_AMT][r]);
      staff[a].conv = _num_(m[C.PAYCONV][r]); // rate 0..1 (last seen for agent in month)
    }
    var conv = t.chats ? (t.orders / t.chats) : 0;

    out.sales_kpi = [
      { label: 'เป้ายอดขาย (บาท)',        value: target ? _fmt_(target) : '—',                 sub: target ? 'เป้ารายเดือน' : 'ยังไม่ได้ตั้งเป้า' },
      { label: 'ยอดขายปัจจุบัน (บาท)',     value: _fmt_(t.sales),                                sub: 'เดือน ' + latest },
      { label: '% ทำยอดขาย',              value: target ? _pct_(t.sales / target * 100) : '—',  sub: 'เทียบเป้ารายเดือน' },
      { label: 'ยอดขาด/เกินเป้า (บาท)',    value: target ? _fmt_(t.sales - target) : '—',        sub: target ? '' : 'ไม่มีเป้าในชีต' },
      { label: 'จำนวนแชท (ไม่ซ้ำ)',        value: _fmt_(t.chats),                                sub: 'ลูกค้าที่ทักเข้ามาทั้งหมด' },
      { label: 'ออเดอร์ที่เกิดขึ้น',        value: _fmt_(t.orders),                               sub: 'คำสั่งซื้อที่ชำระ' },
      { label: 'อัตราแปลงแชทเป็นออเดอร์',   value: _pct_(conv * 100),                             sub: 'เทียบจากแชททั้งหมด' }
    ];

    var over = Math.max(0, t.recv - t.w30);
    var slaTot = t.recv || 1;
    out.sla_response = [
      { label: 'ภายใน 1 นาที',  pct: _pct_(t.w1 / slaTot * 100) },
      { label: 'ภายใน 5 นาที',  pct: _pct_((t.w5 - t.w1) / slaTot * 100) },
      { label: 'ภายใน 10 นาที', pct: _pct_((t.w10 - t.w5) / slaTot * 100) },
      { label: 'ภายใน 30 นาที', pct: _pct_((t.w30 - t.w10) / slaTot * 100) },
      { label: 'เกิน 30 นาที',   pct: _pct_(over / slaTot * 100) }
    ];

    out.sales_funnel = [
      { label: 'แชททั้งหมด (ไม่ซ้ำ)',    value: t.chats },
      { label: 'แชทที่ได้รับการต้อนรับ', value: t.recv },
      { label: 'ชี้นำการสั่งซื้อ',        value: t.guided },
      { label: 'ออเดอร์ที่สำเร็จ',        value: t.orders }
    ];

    out.sales_summary = [
      { label: 'ยอดขายรวม (บาท)',            value: _fmt_(t.sales) },
      { label: 'ออเดอร์รวม',                 value: _fmt_(t.orders) },
      { label: 'จำนวนแชทรวม (ไม่ซ้ำ)',       value: _fmt_(t.chats) },
      { label: 'ลูกค้าไม่พอใจ (รีวิวเชิงลบ)', value: _fmt_(t.neg) }
    ];

    var staffArr = Object.keys(staff).map(function (k) { return staff[k]; });
    staffArr.sort(function (a, b) { return b.sales - a.sales; });
    out.staff_summary = staffArr.map(function (s) {
      return { name: s.name, chats: _fmt_(s.chats), close_rate: _pct_(s.conv * 100), sales: _fmt_(s.sales) };
    });

    out.sales_trend = months.map(function (mm) {
      return { period: 'month', label: _thMon_(mm), sales: Math.round(trendByMonth[mm]), goal: target || 0 };
    });

    out.__latest = latest; // for reference
  }

  /* ===== Detail_ALL (per chat) -> top shops + close grid ===== */
  var xS = ss.getSheetByName('Detail_ALL');
  if (xS && out.__latest) {
    var X = _cols_(xS, [C.D_MONTH, C.D_SHOP, C.D_PAID]);
    var xn = X.n, xm = X.col, latest2 = out.__latest;

    // top shops by chats (latest month)
    var chatByShop = {}, paidByShop = {}, totLatest = 0;
    // per-shop per-month for close grid (all months)
    var gridChats = {}, gridPaid = {}, shopTotal = {}, monthsSet = [];
    for (var j = 0; j < xn; j++) {
      var mo3 = xm[C.D_MONTH][j], shop = xm[C.D_SHOP][j] || '-', paid = _yes_(xm[C.D_PAID][j]);
      if (monthsSet.indexOf(mo3) < 0) monthsSet.push(mo3);
      var key = shop + '||' + mo3;
      gridChats[key] = (gridChats[key] || 0) + 1;
      if (paid) gridPaid[key] = (gridPaid[key] || 0) + 1;
      shopTotal[shop] = (shopTotal[shop] || 0) + 1;
      if (mo3 === latest2) {
        totLatest++;
        chatByShop[shop] = (chatByShop[shop] || 0) + 1;
        if (paid) paidByShop[shop] = (paidByShop[shop] || 0) + 1;
      }
    }
    monthsSet.sort();

    var shopsByChat = Object.keys(chatByShop).sort(function (a, b) { return chatByShop[b] - chatByShop[a]; });
    out.top_chat = shopsByChat.slice(0, 5).map(function (s) {
      return { name: s, chats: chatByShop[s], pct: _pct_(chatByShop[s] / (totLatest || 1) * 100) };
    });
    var totPaid = 0; Object.keys(paidByShop).forEach(function (k) { totPaid += paidByShop[k]; });
    var shopsByPaid = Object.keys(paidByShop).sort(function (a, b) { return paidByShop[b] - paidByShop[a]; });
    out.top_sales = shopsByPaid.slice(0, 5).map(function (s) {
      return { name: s, sales: paidByShop[s], pct: _pct_(paidByShop[s] / (totPaid || 1) * 100) };
    });

    // close grid: top 5 shops (by total chats) x all months
    var topShops = Object.keys(shopTotal).sort(function (a, b) { return shopTotal[b] - shopTotal[a]; }).slice(0, 5);
    var grid = [];
    topShops.forEach(function (s) {
      monthsSet.forEach(function (mm) {
        var kk = s + '||' + mm, ch = gridChats[kk] || 0, pd = gridPaid[kk] || 0;
        grid.push({ store: s, month: _thMon_(mm), pct: _pct_(ch ? pd / ch * 100 : 0) });
      });
    });
    out.store_close_grid = grid;
  }

  delete out.__latest;
  return out;
}

function doGet(e) {
  try {
    var token = (e && e.parameter && e.parameter.token) || '';
    if (!_checkToken_(token)) return _json_({ ok: false, error: 'unauthorized' });
    var ss = _ss_();
    if (!ss) return _json_({ ok: false, error: 'no spreadsheet' });
    return _json_({ ok: true, data: _buildDashboard_(ss) });
  } catch (err) {
    return _json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    if (!_checkToken_(body.token || '')) return _json_({ ok: false, error: 'unauthorized' });
    var ss = _ss_();
    var sheet = ss.getSheetByName(body.sheet);
    if (!sheet) return _json_({ ok: false, error: 'sheet not found: ' + body.sheet });
    var values = sheet.getDataRange().getValues();
    var headers = (values[0] || []).map(function (h) { return String(h).trim(); });
    var colIndex = {}; headers.forEach(function (h, i) { colIndex[h] = i; });
    function rowMatches(rv, mt) {
      return Object.keys(mt || {}).every(function (k) {
        return colIndex[k] !== undefined && String(rv[colIndex[k]]) === String(mt[k]);
      });
    }
    if (body.action === 'create') {
      var nr = headers.map(function (h) { return (body.row && body.row[h] !== undefined) ? body.row[h] : ''; });
      sheet.appendRow(nr); return _json_({ ok: true, action: 'create' });
    }
    if (body.action === 'update') {
      for (var r = 1; r < values.length; r++) if (rowMatches(values[r], body.match)) {
        headers.forEach(function (h, c) { if (body.row && body.row[h] !== undefined) sheet.getRange(r + 1, c + 1).setValue(body.row[h]); });
        return _json_({ ok: true, action: 'update', rowNumber: r + 1 });
      }
      return _json_({ ok: false, error: 'no matching row' });
    }
    if (body.action === 'delete') {
      for (var r2 = 1; r2 < values.length; r2++) if (rowMatches(values[r2], body.match)) {
        sheet.deleteRow(r2 + 1); return _json_({ ok: true, action: 'delete', rowNumber: r2 + 1 });
      }
      return _json_({ ok: false, error: 'no matching row' });
    }
    return _json_({ ok: false, error: 'unknown action: ' + body.action });
  } catch (err) {
    return _json_({ ok: false, error: String(err) });
  }
}
