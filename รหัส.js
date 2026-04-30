// ============================================================
//  POS SYSTEM — ร้านลาบบ้านสวน
//  Professional Restaurant POS Backend
//  Google Apps Script + Google Sheets
// ============================================================

const SHEET_PRODUCTS = 'Products';
const SHEET_ORDERS = 'Orders';
const SHEET_ITEMS = 'OrderItems';
const SHEET_SETTINGS = 'Settings';
const SHEET_RESERVATIONS = 'Reservations';
const SHEET_STAFF = 'Staff';

// ── Entry Point ──────────────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'pos';

  if (page === 'kitchen') {
    return HtmlService.createTemplateFromFile('kitchen')
      .evaluate()
      .setTitle('ครัว — ลาบบ้านสวน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'menu') {
    var tmpl = HtmlService.createTemplateFromFile('menu');
    tmpl.tableNo = (e && e.parameter && e.parameter.table) || '-';
    return tmpl.evaluate()
      .setTitle('เมนู — ลาบบ้านสวน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('POS — ลาบบ้านสวน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Get Web App URL for QR code generation ──
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ── Google Sheet ID (fallback) ────────────────────────────────
const SPREADSHEET_ID = '1V0j6xcnjrQ2XWy3MGHYM5NbX9yS14iHJn6Njm_LS1OY';

// ── Sheet Helpers ─────────────────────────────────────────────
function getSS_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) { }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet_(name) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function sheetToObjects_(sheetName, headers) {
  const sh = getSheet_(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const hdr = headers || data[0];
  return data.slice(1).map((r, idx) => {
    const obj = { _row: idx + 2 };
    hdr.forEach((h, i) => { obj[h] = r[i]; });
    return obj;
  });
}

// ── Initial Setup ─────────────────────────────────────────────
function setupSheets() {
  const ss = getSS_();
  const headerStyle = (sh, cols) => {
    sh.getRange(1, 1, 1, cols).setFontWeight('bold')
      .setBackground('#FF6B35').setFontColor('#fff');
    sh.setFrozenRows(1);
  };

  // ─── Products (สร้างครั้งแรกเท่านั้น ไม่ล้างของเดิม) ───
  let p = ss.getSheetByName(SHEET_PRODUCTS);
  if (!p) p = ss.insertSheet(SHEET_PRODUCTS);
  if (!p.getRange('A1').getValue()) {
    p.appendRow(['id', 'name', 'category', 'price', 'stock', 'unit', 'active',
      'emoji', 'spice_default', 'options', 'cost', 'sort_order']);
    headerStyle(p, 12);

    const menu = [
      // ─── 🌟 เมนูแนะนำ ───
      ['M001', 'ลาบหมูคั่ว', '🌟 แนะนำ', 80, 99, 'จาน', true, '🥩', 'กลาง', 'เพิ่มไข่มดแดง:20', 35, 1],
      ['M002', 'ต้มแซ่บกระดูกอ่อน', '🌟 แนะนำ', 90, 99, 'ชาม', true, '🍲', 'กลาง', '', 40, 2],
      ['M003', 'คอหมูย่าง', '🌟 แนะนำ', 100, 99, 'จาน', true, '🔥', '', '', 45, 3],

      // ─── ลาบ / น้ำตก ───
      ['M004', 'ลาบหมู', 'ลาบ/น้ำตก', 70, 99, 'จาน', true, '🥩', 'กลาง', 'เพิ่มไข่มดแดง:20', 30, 10],
      ['M005', 'ลาบไก่', 'ลาบ/น้ำตก', 70, 99, 'จาน', true, '🍗', 'กลาง', 'เพิ่มไข่มดแดง:20', 30, 11],
      ['M006', 'ลาบเนื้อ', 'ลาบ/น้ำตก', 90, 99, 'จาน', true, '🥩', 'กลาง', 'เพิ่มไข่มดแดง:20', 40, 12],
      ['M007', 'น้ำตกหมู', 'ลาบ/น้ำตก', 80, 99, 'จาน', true, '🥩', 'กลาง', '', 35, 13],
      ['M008', 'น้ำตกเนื้อ', 'ลาบ/น้ำตก', 100, 99, 'จาน', true, '🥩', 'กลาง', '', 45, 14],

      // ─── ต้มแซ่บ ───
      ['M009', 'ต้มแซ่บไก่บ้าน', 'ต้มแซ่บ', 100, 99, 'ชาม', true, '🍲', 'กลาง', '', 45, 20],
      ['M010', 'ต้มแซ่บเนื้อ', 'ต้มแซ่บ', 120, 99, 'ชาม', true, '🍲', 'กลาง', '', 55, 21],
      ['M011', 'อ่อมหมู', 'ต้มแซ่บ', 90, 99, 'ชาม', true, '🍲', 'กลาง', '', 40, 22],
      ['M012', 'อ่อมเนื้อ', 'ต้มแซ่บ', 90, 99, 'ชาม', true, '🍲', 'กลาง', '', 40, 23],

      // ─── เมนูย่าง ───
      ['M013', 'ไก่ย่างบ้านสวน', 'เมนูย่าง', 120, 99, 'ตัว', true, '🍗', '', '', 50, 30],
      ['M014', 'เสือร้องไห้', 'เมนูย่าง', 120, 99, 'จาน', true, '🥩', '', '', 55, 31],

      // ─── เครื่องเคียง ───
      ['M015', 'ข้าวเหนียว', 'เครื่องเคียง', 10, 999, 'ห่อ', true, '🍚', '', '', 3, 40],
      ['M016', 'ผักสด/ผักลวก', 'เครื่องเคียง', 0, 999, 'จาน', true, '🥬', '', '', 5, 41],
      ['M017', 'แจ่วปลาร้า', 'เครื่องเคียง', 15, 999, 'ถ้วย', true, '🫙', '', '', 5, 42],
      ['M018', 'แจ่วบอง', 'เครื่องเคียง', 15, 999, 'ถ้วย', true, '🫙', '', '', 5, 43],

      // ─── เครื่องดื่ม ───
      ['M019', 'น้ำเปล่า', 'เครื่องดื่ม', 15, 200, 'ขวด', true, '💧', '', '', 7, 50],
      ['M020', 'น้ำอัดลม', 'เครื่องดื่ม', 20, 200, 'กระป๋อง', true, '🥤', '', '', 10, 51],
      ['M021', 'ชาเย็น', 'เครื่องดื่ม', 40, 99, 'แก้ว', true, '🧋', '', '', 15, 52],
      ['M022', 'กาแฟเย็น', 'เครื่องดื่ม', 45, 99, 'แก้ว', true, '☕', '', '', 18, 53],
      ['M023', 'น้ำมะนาว', 'เครื่องดื่ม', 35, 99, 'แก้ว', true, '🍋', '', '', 12, 54],
      ['M024', 'โอเลี้ยง', 'เครื่องดื่ม', 35, 99, 'แก้ว', true, '☕', '', '', 12, 55],
    ];
    menu.forEach(r => p.appendRow(r));
  }

  // ─── Orders ───
  let o = ss.getSheetByName(SHEET_ORDERS);
  if (!o) o = ss.insertSheet(SHEET_ORDERS);
  if (!o.getRange('A1').getValue()) {
    o.appendRow(['order_id', 'date', 'time', 'table_no', 'order_type',
      'total', 'discount', 'discount_type', 'grand_total',
      'payment', 'note', 'cashier', 'status', 'customer_count']);
    headerStyle(o, 14);
  }
  // Always fix header row to correct format
  const correctHeaders = ['order_id', 'date', 'time', 'table_no', 'order_type',
    'total', 'discount', 'discount_type', 'grand_total',
    'payment', 'note', 'cashier', 'status', 'customer_count'];
  o.getRange(1, 1, 1, 14).setValues([correctHeaders]);

  // ─── OrderItems ───
  let i = ss.getSheetByName(SHEET_ITEMS);
  if (!i) i = ss.insertSheet(SHEET_ITEMS);
  if (!i.getRange('A1').getValue()) {
    i.appendRow(['order_id', 'product_id', 'product_name', 'qty',
      'unit_price', 'subtotal', 'spice_level', 'options', 'note', 'item_status']);
    headerStyle(i, 10);
  }
  // Auto-fix OrderItems header
  i.getRange(1, 1, 1, 10).setValues([['order_id', 'product_id', 'product_name', 'qty', 'unit_price', 'subtotal', 'spice_level', 'options', 'note', 'item_status']]);

  // ─── Settings ───
  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) s = ss.insertSheet(SHEET_SETTINGS);
  if (!s.getRange('A1').getValue()) {
    s.appendRow(['key', 'value']);
    headerStyle(s, 2);
    const defaults = [
      ['shop_name', 'ลาบบ้านสวน'],
      ['shop_phone', '090-123-4567'],
      ['vat', '0'],
      ['last_order_num', '0'],
      ['table_count', '10'],
      ['promptpay_id', '0901234567'],
      ['promo_family_price', '299'],
      ['promo_student_discount', '10'],
      ['promo_checkin_enabled', 'true'],
    ];
    defaults.forEach(r => s.appendRow(r));
  }

  // ─── Reservations ───
  let r = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!r) r = ss.insertSheet(SHEET_RESERVATIONS);
  if (!r.getRange('A1').getValue()) {
    r.appendRow(['reservation_id', 'date', 'time', 'table_no', 'customer_name', 'party_size', 'contact', 'note', 'status', 'created_by', 'created_at']);
    headerStyle(r, 11);
  }

  // ─── Staff ───
  let st = ss.getSheetByName(SHEET_STAFF);
  if (!st) st = ss.insertSheet(SHEET_STAFF);
  if (!st.getRange('A1').getValue()) {
    st.appendRow(['name', 'pin', 'role', 'active']);
    headerStyle(st, 4);
    st.appendRow(['พนักงาน 1', '1234', 'cashier', true]);
    st.appendRow(['พนักงาน 2', '5678', 'cashier', true]);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Setup สำเร็จ!' };
}

// ── Staff / Login ────────────────────────────────────────────
function loginWithPin(pin) {
  let staff = sheetToObjects_(SHEET_STAFF, ['name', 'pin', 'role', 'active']);
  // Auto-setup if no staff exists
  if (!staff.length) {
    setupSheets();
    staff = sheetToObjects_(SHEET_STAFF, ['name', 'pin', 'role', 'active']);
  }
  const found = staff.find(s => String(s.pin) === String(pin) && s.active !== false);
  if (found) return { success: true, name: found.name, role: found.role };
  return { success: false };
}

function getStaffList() {
  return sheetToObjects_(SHEET_STAFF, ['name', 'pin', 'role', 'active'])
    .filter(s => s.active !== false)
    .map(s => ({ name: s.name, role: s.role }));
}

// ── Settings ─────────────────────────────────────────────────
function getSettings() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('app_settings');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* fall through */ }
  }
  var sh = getSheet_(SHEET_SETTINGS);
  var data = sh.getDataRange().getValues();
  var s = {};
  for (var i = 1; i < data.length; i++) { s[data[i][0]] = data[i][1]; }
  try { cache.put('app_settings', JSON.stringify(s), 300); } catch(e) { /* cache full */ }
  return s;
}

function saveSetting(key, value) {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

function saveMultipleSettings(kvPairs) {
  kvPairs.forEach(([k, v]) => saveSetting(k, v));
  return { success: true };
}

// ── Products ─────────────────────────────────────────────────
function getProducts() {
  const sh = getSheet_(SHEET_PRODUCTS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .filter(r => r[6] !== false && r[6] !== 'FALSE')
    .map(r => ({
      id: r[0], name: r[1], category: r[2],
      price: Number(r[3]), stock: Number(r[4]),
      unit: r[5] || 'จาน', active: r[6],
      emoji: r[7] || '🍽️', spice_default: r[8] || '',
      options: r[9] || '', cost: Number(r[10]) || 0,
      sort_order: Number(r[11]) || 99,
      image_url: r[12] || ''
    }))
    .sort((a, b) => a.sort_order - b.sort_order);
}

function resolveDriveId_(url) {
  if (!url) return '';
  var s = String(url).trim();
  if (s.indexOf('drive:') === 0) return s.replace('drive:', '');
  var match = s.match(/[?&]id=([^&]+)/);
  if (match) return match[1];
  match = s.match(/\/d\/([^/]+)/);
  if (match) return match[1];
  return s;
}

function getProductsWithImages() {
  return getProducts().map(function(p) {
    var imgData = '';
    if (p.image_url) {
      try {
        var fid = resolveDriveId_(p.image_url);
        if (fid) imgData = getCachedImageData_(fid);
      } catch (e) {}
    }
    return {
      id: p.id, name: p.name, category: p.category,
      price: p.price, stock: p.stock, unit: p.unit, active: p.active,
      emoji: p.emoji, spice_default: p.spice_default, options: p.options,
      cost: p.cost, sort_order: p.sort_order,
      image_data: imgData,
      image_url: p.image_url
    };
  });
}

// Customer menu - fast load WITHOUT images
function getMenuForCustomer() {
  var prods = getProducts().map(function(p) {
    return {
      id: p.id, name: p.name, category: p.category,
      price: p.price, unit: p.unit, emoji: p.emoji,
      spice_default: p.spice_default, options: p.options,
      image_url: p.image_url || ''
    };
  });
  var sett = getSettings();
  return { products: prods, shopName: sett.shop_name || 'ลาบบ้านสวน' };
}

// Fetch a single product image (with 6hr cache)
function getProductImage(productId) {
  var prods = getProducts();
  for (var i = 0; i < prods.length; i++) {
    if (String(prods[i].id) === String(productId) && prods[i].image_url) {
      var fid = resolveDriveId_(prods[i].image_url);
      if (fid) return getCachedImageData_(fid);
    }
  }
  return '';
}


function getCategories() {
  return [...new Set(getProducts().map(p => p.category))];
}

function addProduct(data) {
  const sh = getSheet_(SHEET_PRODUCTS);
  const id = 'M' + String(Date.now()).slice(-6);
  sh.appendRow([id, data.name, data.category, Number(data.price),
    Number(data.stock), data.unit || 'จาน', true,
    data.emoji || '🍽️', data.spice_default || '', data.options || '',
    Number(data.cost) || 0, Number(data.sort_order) || 99]);
  return { success: true, id };
}

function updateProduct(data) {
  const sh = getSheet_(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sh.getRange(i + 1, 2, 1, 11).setValues([[
        data.name, data.category, Number(data.price), Number(data.stock),
        data.unit, data.active !== false, data.emoji || '🍽️',
        data.spice_default || '', data.options || '',
        Number(data.cost) || 0, Number(data.sort_order) || 99
      ]]);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบสินค้า' };
}

function deleteProduct(id) {
  const sh = getSheet_(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { sh.getRange(i + 1, 7).setValue(false); return { success: true }; }
  }
  return { success: false };
}

function updateStock(id, delta) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_(SHEET_PRODUCTS);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        var current = Number(rows[i][4]) || 0;
        if (delta < 0 && current + delta < 0) {
          return { error: 'stock_insufficient', current: current };
        }
        var ns = Math.max(0, current + delta);
        sh.getRange(i + 1, 5).setValue(ns);
        SpreadsheetApp.flush();
        return ns;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Orders ───────────────────────────────────────────────────
function generateOrderId_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_(SHEET_SETTINGS);
    var data = sh.getDataRange().getValues();
    var num = 0, row = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'last_order_num') { num = Number(data[i][1]) || 0; row = i + 1; break; }
    }
    num++;
    if (row > 0) {
      sh.getRange(row, 2).setValue(num);
    } else {
      sh.appendRow(['last_order_num', num]);
    }
    SpreadsheetApp.flush();
    var d = new Date();
    var prefix = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyyMMdd');
    return 'ORD-' + prefix + '-' + String(num).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

function saveOrder(orderData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var orderId = generateOrderId_();
    var now = new Date();
    var date = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
    var time = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');
    var orderStatus = orderData.initialStatus || 'cooking';
    var itemStatus = orderStatus === 'completed' ? 'served' : 'cooking';

    // --- Pre-validate stock ---
    var pSh = getSheet_(SHEET_PRODUCTS);
    var pData = pSh.getDataRange().getValues();
    var stockMap = {};
    for (var p = 1; p < pData.length; p++) {
      stockMap[String(pData[p][0])] = { row: p + 1, stock: Number(pData[p][4]) || 0 };
    }
    for (var v = 0; v < orderData.items.length; v++) {
      var si = stockMap[String(orderData.items[v].id)];
      if (si && si.stock < orderData.items[v].qty) {
        return { success: false, message: 'สินค้า ' + orderData.items[v].name + ' เหลือ ' + si.stock + ' ไม่พอ' };
      }
    }

    // --- Write order (single appendRow) ---
    var oSh = getSheet_(SHEET_ORDERS);
    oSh.appendRow([
      orderId, date, time,
      orderData.tableNo || '-',
      orderData.orderType || 'dine-in',
      orderData.total,
      orderData.discount || 0,
      orderData.discountType || '',
      orderData.grandTotal,
      orderData.payment || '',
      orderData.note || '',
      orderData.cashier || '',
      orderStatus,
      orderData.customerCount || 1
    ]);
    oSh.getRange(oSh.getLastRow(), 1).setNumberFormat('@');

    // --- Batch write items (single setValues) ---
    var iSh = getSheet_(SHEET_ITEMS);
    var itemRows = [];
    for (var j = 0; j < orderData.items.length; j++) {
      var item = orderData.items[j];
      itemRows.push([
        orderId, item.id, item.name, item.qty,
        item.price, item.qty * item.price,
        item.spice || '', item.options || '', item.note || '', itemStatus
      ]);
    }
    if (itemRows.length > 0) {
      var startRow = iSh.getLastRow() + 1;
      iSh.getRange(startRow, 1, itemRows.length, 10).setValues(itemRows);
      // Set orderId column to text format
      iSh.getRange(startRow, 1, itemRows.length, 1).setNumberFormat('@');
    }

    // --- Batch update stock ---
    for (var k = 0; k < orderData.items.length; k++) {
      var sk = stockMap[String(orderData.items[k].id)];
      if (sk) {
        var ns = Math.max(0, sk.stock - orderData.items[k].qty);
        pSh.getRange(sk.row, 5).setValue(ns);
      }
    }

    SpreadsheetApp.flush();

    var settings = getSettings();
    return {
      success: true, orderId: orderId,
      shopName: settings.shop_name || 'ลาบบ้านสวน',
      shopPhone: settings.shop_phone || '',
      date: date, time: time
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateOrderStatus(orderId, status, tableNo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_(SHEET_ORDERS);
    var data = sh.getDataRange().getValues();
    var oid = String(orderId).trim();
    var tbl = tableNo ? String(tableNo).trim() : '';
    for (var i = data.length - 1; i >= 1; i--) {
      var matchId = String(data[i][0]).trim() === oid;
      var matchTbl = !tbl || String(data[i][3]).trim().replace(/['\"]/g, '') === tbl;
      if (matchId && matchTbl) {
        var oldStatus = String(data[i][12]);
        sh.getRange(i + 1, 13).setValue(status);

        if (status === 'cooking') {
          var iSh = getSheet_(SHEET_ITEMS);
          var iData = iSh.getDataRange().getValues();
          for (var j = 1; j < iData.length; j++) {
            if (String(iData[j][0]).trim() === oid) {
              iSh.getRange(j + 1, 10).setValue('cooking');
            }
          }
        } else if (status === 'served') {
          var iSh2 = getSheet_(SHEET_ITEMS);
          var iData2 = iSh2.getDataRange().getValues();
          for (var j2 = 1; j2 < iData2.length; j2++) {
            if (String(iData2[j2][0]).trim() === oid && String(iData2[j2][9]) !== 'served') {
              iSh2.getRange(j2 + 1, 10).setValue('served');
            }
          }
        } else if (status === 'completed') {
          var tblStr = tbl || String(data[i][3]).trim().replace(/['\"]/g, '');
          try {
            var nSh = getSheet_('Notifications');
            var nData = nSh.getDataRange().getValues();
            for (var k = nData.length - 1; k >= 1; k--) {
              if (String(nData[k][2]).trim() === tblStr && String(nData[k][4]) !== 'done') {
                nSh.getRange(k + 1, 5).setValue('done');
              }
            }
          } catch (ne) { /* Notifications sheet may not exist */ }
        }

        SpreadsheetApp.flush();
        return { success: true, oldStatus: oldStatus, newStatus: status, row: i + 1, table: String(data[i][3]) };
      }
    }
    return { success: false, message: 'ไม่พบ Order: ' + oid + ' table: ' + tbl };
  } finally {
    lock.releaseLock();
  }
}

// Complete ALL active orders for a table at once
function completeTable(tableNo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    var sh = getSheet_(SHEET_ORDERS);
    var data = sh.getDataRange().getValues();
    var tblStr = String(tableNo).trim();
    var updated = 0;
    
    for (var i = data.length - 1; i >= 1; i--) {
      var rowDate = data[i][1];
      if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
      var rowTbl = String(data[i][3]).trim().replace(/['"]/g, '');
      var rowStatus = String(data[i][12]);
      
      if (String(rowDate) === today && rowTbl === tblStr && ['new', 'cooking', 'served'].indexOf(rowStatus) !== -1) {
        sh.getRange(i + 1, 13).setValue('completed');
        updated++;
      }
    }
    
    // Clear notifications for this table
    try {
      var nSh = getSheet_('Notifications');
      var nData = nSh.getDataRange().getValues();
      for (var k = nData.length - 1; k >= 1; k--) {
        if (String(nData[k][2]).trim() === tblStr && String(nData[k][4]) !== 'done') {
          nSh.getRange(k + 1, 5).setValue('done');
        }
      }
    } catch(ne) {}
    
    SpreadsheetApp.flush();
    return { success: true, updatedOrders: updated, table: tblStr };
  } finally {
    lock.releaseLock();
  }
}

function cancelOrder(orderId, reason) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_(SHEET_ORDERS);
    var data = sh.getDataRange().getValues();
    var oid = String(orderId).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === oid) {
        sh.getRange(i + 1, 13).setValue('cancelled');
        var note = data[i][10] ? data[i][10] + ' | ยกเลิก: ' + reason : 'ยกเลิก: ' + reason;
        sh.getRange(i + 1, 11).setValue(note);
        var items = getOrderItems_(orderId);
        for (var x = 0; x < items.length; x++) {
          updateStock(items[x].product_id, items[x].qty);
        }
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false };
  } finally {
    lock.releaseLock();
  }
}

function getOrderItems_(orderId) {
  var sh = getSheet_(SHEET_ITEMS);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var oid = String(orderId).trim();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === oid) {
      results.push({
        product_id: data[i][1], product_name: data[i][2], qty: Number(data[i][3]),
        unit_price: Number(data[i][4]), subtotal: Number(data[i][5]),
        spice: data[i][6], options: data[i][7], note: data[i][8],
        item_status: String(data[i][9] || 'cooking'),
        row: i + 1
      });
    }
  }
  return results;
}

function updateItemStatus(orderId, itemRow, newStatus) {
  var iSh = getSheet_(SHEET_ITEMS);
  iSh.getRange(itemRow, 10).setValue(newStatus);
  var data = iSh.getDataRange().getValues();
  var allServed = true;
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(orderId).trim()) {
      found = true;
      if (String(data[i][9]) !== 'served') allServed = false;
    }
  }
  if (!found) allServed = false;
  if (allServed) {
    var oSh = getSheet_(SHEET_ORDERS);
    var oData = oSh.getDataRange().getValues();
    for (var j = oData.length - 1; j >= 1; j--) {
      if (String(oData[j][0]).trim() === String(orderId).trim()) {
        oSh.getRange(j + 1, 13).setValue('served');
        break;
      }
    }
  }
  SpreadsheetApp.flush();
  return { success: true, allServed: allServed };
}

// ── Customer Order Tracking ──────────────────────────────────
function trackOrder(orderId, tableNo) {
  try {
    var oSh = getSheet_(SHEET_ORDERS);
    var oData = oSh.getDataRange().getValues();
    var oid = orderId ? String(orderId).trim() : '';
    var tbl = tableNo ? String(tableNo).trim() : '';
    var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    var order = null;
    var debugInfo = [];

    if (oid) {
      for (var i = oData.length - 1; i >= 1; i--) {
        if (String(oData[i][0]).trim() === oid) {
          order = { orderId: String(oData[i][0]), date: oData[i][1], time: oData[i][2], tableNo: String(oData[i][3]), status: String(oData[i][12] || 'new'), total: Number(oData[i][8]) };
          break;
        }
      }
    }

    if (!order && tbl) {
      for (var j = oData.length - 1; j >= 1; j--) {
        var rowTbl = String(oData[j][3]).trim().replace(/['"]/g, '');
        var rowStatus = String(oData[j][12] || '').trim();
        var rowDate = oData[j][1] instanceof Date ? Utilities.formatDate(oData[j][1], 'Asia/Bangkok', 'yyyy-MM-dd') : String(oData[j][1]).trim();
        if (rowTbl == tbl && rowDate === today && debugInfo.length < 5) {
          debugInfo.push('r' + (j + 1) + ':id=' + String(oData[j][0]).substring(0, 20) + ',st=' + rowStatus);
        }
        if (rowTbl == tbl && rowDate === today && ['new', 'cooking', 'served', 'completed'].indexOf(rowStatus) !== -1) {
          order = { orderId: String(oData[j][0]), date: rowDate, time: oData[j][2], tableNo: String(oData[j][3]), status: rowStatus, total: Number(oData[j][8]) };
          break;
        }
      }
    }

    if (!order) {
      for (var k = Math.max(1, oData.length - 5); k < oData.length; k++) {
        debugInfo.push('r' + (k + 1) + ':id=[' + String(oData[k][0]).substring(0, 25) + ']|t=' + typeof oData[k][0] + '|st=' + String(oData[k][12]));
      }
      return { success: false, message: 'SEARCH:oid=[' + oid + ']tbl=[' + tbl + ']today=' + today + '|ROWS=' + oData.length + '|' + debugInfo.join('|') };
    }

    var rawItems = getOrderItems_(order.orderId);
    var items = [];
    for (var m = 0; m < rawItems.length; m++) {
      items.push({ name: rawItems[m].product_name, qty: rawItems[m].qty, price: rawItems[m].unit_price, subtotal: rawItems[m].subtotal, spice: rawItems[m].spice, note: rawItems[m].note, item_status: rawItems[m].item_status || 'cooking' });
    }
    var servedCount = 0;
    for (var n = 0; n < items.length; n++) { if (items[n].item_status === 'served') servedCount++; }
    var allServed = items.length > 0 && servedCount === items.length;

    return { success: true, order: order, items: items, allServed: allServed, servedCount: servedCount };
  } catch (e) {
    return { success: false, message: 'ERR:' + String(e) };
  }
}

function getOrderTracking(orderId) {
  try {
    var oSh = getSheet_(SHEET_ORDERS);
    var oData = oSh.getDataRange().getValues();
    var oid = orderId ? String(orderId).trim() : '';
    var order = null;
    for (var i = oData.length - 1; i >= 1; i--) {
      if (String(oData[i][0]).trim() === oid) {
        order = { orderId: String(oData[i][0]), date: String(oData[i][1]), time: String(oData[i][2]), tableNo: String(oData[i][3]), status: String(oData[i][12] || 'new'), total: Number(oData[i][8]) };
        break;
      }
    }
    if (!order) return { success: false, message: 'ไม่พบออเดอร์' };
    var items = getOrderItems_(String(order.orderId));
    var mapped = [];
    for (var m = 0; m < items.length; m++) {
      mapped.push({ name: items[m].product_name, qty: items[m].qty, price: items[m].unit_price, subtotal: items[m].subtotal, spice: items[m].spice, note: items[m].note, item_status: items[m].item_status || 'cooking' });
    }
    var sc = 0;
    for (var n = 0; n < mapped.length; n++) { if (mapped[n].item_status === 'served') sc++; }
    return { success: true, order: order, items: mapped, allServed: mapped.length > 0 && sc === mapped.length, servedCount: sc };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}

function getOrderTrackingByTable(tableNo) {
  try {
    var oSh = getSheet_(SHEET_ORDERS);
    var oData = oSh.getDataRange().getValues();
    var tbl = tableNo ? String(tableNo).trim() : '';
    var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    var order = null;
    for (var j = oData.length - 1; j >= 1; j--) {
      var rowTbl = String(oData[j][3]).trim().replace(/['"]/g, '');
      var rowSt = String(oData[j][12] || '').trim();
      var rowDt = oData[j][1] instanceof Date ? Utilities.formatDate(oData[j][1], 'Asia/Bangkok', 'yyyy-MM-dd') : String(oData[j][1]).trim();
      if (rowTbl == tbl && rowDt === today && ['new', 'cooking', 'served', 'completed'].indexOf(rowSt) !== -1) {
        order = { orderId: String(oData[j][0]), date: rowDt, time: String(oData[j][2]), tableNo: String(oData[j][3]), status: rowSt, total: Number(oData[j][8]) };
        break;
      }
    }
    if (!order) return { success: false, message: 'ไม่พบออเดอร์สำหรับโต๊ะนี้' };
    var items = getOrderItems_(String(order.orderId));
    var mapped = [];
    for (var m = 0; m < items.length; m++) {
      mapped.push({ name: items[m].product_name, qty: items[m].qty, price: items[m].unit_price, subtotal: items[m].subtotal, spice: items[m].spice, note: items[m].note, item_status: items[m].item_status || 'cooking' });
    }
    var sc = 0;
    for (var n = 0; n < mapped.length; n++) { if (mapped[n].item_status === 'served') sc++; }
    return { success: true, order: order, items: mapped, allServed: mapped.length > 0 && sc === mapped.length, servedCount: sc };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}


// Customer calls staff
function callStaff(tableNo, type) {
  const sh = getSheet_('Notifications');
  const now = new Date();
  const time = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');
  const date = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
  sh.appendRow([date, time, tableNo, type, 'pending']);
  return { success: true, time: time };
}

// Get pending notifications for POS
function getNotifications() {
  var sh = getSheet_('Notifications');
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(data[i][0]).trim();
    var rowStatus = String(data[i][4] || '').trim();
    if (rowDate === today && rowStatus === 'pending') {
      results.push({
        row: i + 1,
        date: rowDate,
        time: String(data[i][1]),
        tableNo: String(data[i][2]),
        type: String(data[i][3]),
        status: rowStatus
      });
    }
  }
  return results;
}

// Dismiss notification
function dismissNotification(row) {
  const sh = getSheet_('Notifications');
  sh.getRange(row, 5).setValue('done');
  return { success: true };
}
function getKitchenOrders() {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const activeStatuses = ['new', 'cooking'];

  const orders = data.slice(1)
    .filter(r => {
      let rowDate = r[1];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
      }
      return String(rowDate) === today && activeStatuses.includes(String(r[12]));
    })
    .map(r => ({
      orderId: r[0], date: r[1], time: r[2],
      tableNo: r[3], orderType: r[4],
      status: String(r[12]),
      items: getOrderItems_(r[0])
    }));

  return orders.sort((a, b) => a.time > b.time ? 1 : -1);
}

// ── Order History ────────────────────────────────────────────
function getOrderHistory(days) {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - (days || 7));
  const cutoffStr = Utilities.formatDate(cutoff, 'Asia/Bangkok', 'yyyy-MM-dd');

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rowDate = r[1];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    } else {
      rowDate = String(rowDate).trim();
    }
    if (rowDate >= cutoffStr) {
      results.push({
        orderId: String(r[0]), date: rowDate, time: String(r[2]),
        tableNo: String(r[3]), orderType: String(r[4]),
        total: Number(r[5]) || 0, discount: Number(r[6]) || 0,
        discountType: String(r[7] || ''), grandTotal: Number(r[8]) || 0,
        payment: String(r[9] || ''), note: String(r[10] || ''), cashier: String(r[11] || ''),
        status: String(r[12] || ''), customerCount: Number(r[13]) || 1
      });
    }
  }
  results.reverse();
  return results;
}

function getOrderDetail(orderId) {
  const orders = getOrderHistory(30);
  const order = orders.find(o => o.orderId === orderId);
  if (!order) return null;
  order.items = getOrderItems_(orderId);
  return order;
}

function generateReservationId_() {
  const now = new Date();
  const prefix = 'R' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd');
  const sh = getSheet_(SHEET_RESERVATIONS);
  const rows = sh.getDataRange().getValues();
  const seq = rows.length;
  return prefix + '-' + String(seq).padStart(3, '0');
}

function saveReservation(data) {
  const sh = getSheet_(SHEET_RESERVATIONS);
  const reservationId = generateReservationId_();
  const now = new Date();
  const createdAt = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([
    reservationId,
    data.date || Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
    data.time || '18:00',
    data.tableNo || '-',
    data.customerName || 'Guest',
    Number(data.partySize) || 1,
    data.contact || '',
    data.note || '',
    'reserved',
    data.createdBy || '',
    createdAt
  ]);
  return { success: true, reservationId };
}

function updateReservationStatus(reservationId, status) {
  const sh = getSheet_(SHEET_RESERVATIONS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      sh.getRange(i + 1, 9).setValue(status);
      return { success: true };
    }
  }
  return { success: false };
}

function cancelReservation(reservationId, reason) {
  const sh = getSheet_(SHEET_RESERVATIONS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      sh.getRange(i + 1, 9).setValue('cancelled');
      const note = data[i][7] ? data[i][7] + ' | ยกเลิก: ' + reason : 'ยกเลิก: ' + reason;
      sh.getRange(i + 1, 8).setValue(note);
      return { success: true };
    }
  }
  return { success: false };
}

function getReservations(dateStr) {
  const sh = getSheet_(SHEET_RESERVATIONS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const today = dateStr || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  return data.slice(1)
    .filter(r => String(r[1]) === today && r[8] !== 'cancelled')
    .map(r => ({
      reservationId: r[0], date: r[1], time: r[2], tableNo: r[3],
      customerName: r[4], partySize: Number(r[5]) || 1, contact: r[6],
      note: r[7], status: r[8], createdBy: r[9], createdAt: r[10]
    }));
}

function getItemCounts_(orderId) {
  const sh = getSheet_(SHEET_ITEMS);
  const data = sh.getDataRange().getValues();
  const items = data.slice(1).filter(r => String(r[0]) === orderId);
  const served = items.filter(r => String(r[9]) === 'served').length;
  return { total: items.length, served: served };
}

function getTableStatus() {
  try {
    const settings = getSettings();
    const count = Math.max(1, Number(settings.table_count) || 10);
    const orders = getSheet_(SHEET_ORDERS).getDataRange().getValues();
    const reservations = getReservations();
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

    const tables = [];
    for (let t = 1; t <= count; t++) {
      const tableStr = String(t);
      const todayOrders = orders.slice(1).filter(r => {
        let rowDate = r[1];
        if (rowDate instanceof Date) {
          rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
        }
        return String(rowDate) === today && String(r[3]).trim().replace(/['"]/g, '') === tableStr && ['new', 'cooking', 'served'].includes(String(r[12]));
      });

      if (todayOrders.length > 0) {
        // Aggregate all active orders
        var combinedTotal = 0;
        var combinedItemCount = 0;
        var combinedServed = 0;
        var allOids = [];
        var latestStatus = 'served'; // will degrade to cooking/new
        var firstTime = '';

        todayOrders.forEach(function(r) {
          var oid = String(r[0]);
          allOids.push(oid);
          combinedTotal += Number(r[8]) || 0;
          var info = getItemCounts_(oid);
          combinedItemCount += info.total;
          combinedServed += info.served;
          var st = String(r[12]);
          // Priority: new > cooking > served (show most urgent status)
          if (st === 'new') latestStatus = 'new';
          else if (st === 'cooking' && latestStatus !== 'new') latestStatus = 'cooking';
        });

        // Use latest order for time display, first for orderId
        var latestOrder = todayOrders[todayOrders.length - 1];
        var orderTime = latestOrder[2];
        if (orderTime instanceof Date) {
          orderTime = Utilities.formatDate(orderTime, 'Asia/Bangkok', 'HH:mm');
        } else {
          orderTime = String(orderTime || '');
        }
        var firstOid = String(todayOrders[0][0]);
        var firstTime2 = todayOrders[0][2];
        if (firstTime2 instanceof Date) firstTime2 = Utilities.formatDate(firstTime2, 'Asia/Bangkok', 'HH:mm');
        else firstTime2 = String(firstTime2 || '');

        tables.push({
          no: t,
          status: latestStatus,
          orderId: firstOid,
          allOrderIds: allOids,
          orderCount: todayOrders.length,
          orderType: String(latestOrder[4]),
          total: combinedTotal,
          time: firstTime2,
          itemCount: combinedItemCount,
          servedCount: combinedServed,
          totalItems: combinedItemCount
        });
        continue;
      }

      const reservation = reservations.find(r => String(r.tableNo).trim() === tableStr && ['reserved', 'arrived', 'seated'].includes(r.status));
      if (reservation) {
        tables.push({
          no: t,
          status: String(reservation.status),
          reservationId: String(reservation.reservationId),
          customerName: String(reservation.customerName || ''),
          partySize: Number(reservation.partySize) || 1,
          contact: String(reservation.contact || ''),
          resNote: String(reservation.note || ''),
          total: 0,
          time: String(reservation.time || ''),
          itemCount: 0,
          orderCount: 0
        });
        continue;
      }

      tables.push({
        no: t,
        status: 'available',
        orderId: '',
        total: 0,
        time: '',
        itemCount: 0,
        orderCount: 0
      });
    }
    return tables;
  } catch (e) {
    return [{ no: 1, status: 'available', orderId: '', total: 0, time: '', itemCount: 0, error: String(e) }];
  }
}

function getTableDetail(orderId) {
  if (!orderId) return { items: [] };
  return { items: getOrderItems_(orderId) };
}

function getTableDetailByTableNo(tableNo) {
  try {
    var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    var oSh = getSheet_(SHEET_ORDERS);
    var orders = oSh.getDataRange().getValues();
    var tblStr = String(tableNo).trim();
    
    // Find all active orders for this table today
    var activeOrders = orders.slice(1).filter(function(r) {
      var rowDate = r[1];
      if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
      return String(rowDate) === today && 
             String(r[3]).trim().replace(/['"]/g, '') === tblStr && 
             ['new', 'cooking', 'served'].indexOf(String(r[12])) !== -1;
    });
    
    var result = { orders: [], allItems: [], orderCount: activeOrders.length };
    
    activeOrders.forEach(function(r, idx) {
      var oid = String(r[0]);
      var items = getOrderItems_(oid);
      items.forEach(function(item) {
        item.orderIdx = idx + 1;
        item.fromOrderId = oid;
      });
      result.orders.push({
        orderId: oid,
        orderIdx: idx + 1,
        status: String(r[12]),
        total: Number(r[8]) || 0
      });
      result.allItems = result.allItems.concat(items);
    });
    
    return result;
  } catch(e) {
    return { orders: [], allItems: [], orderCount: 0, error: String(e) };
  }
}

function countOrderItems_(orderId) {
  try {
    const sh = getSheet_(SHEET_ITEMS);
    const data = sh.getDataRange().getValues();
    return data.slice(1).filter(r => String(r[0]) === orderId).length;
  } catch (e) { return 0; }
}

// ── Dashboard ────────────────────────────────────────────────
function getDashboardData() {
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var report = getReport(today, today);
  // Add recent 5 orders
  var recent = getOrderHistory(1).slice(0, 5);
  report.recentOrders = recent;
  return report;
}

// ── Reports ──────────────────────────────────────────────────
function getReport(startDate, endDate) {
  const oSh = getSheet_(SHEET_ORDERS);
  const iSh = getSheet_(SHEET_ITEMS);
  const orders = oSh.getDataRange().getValues().slice(1);
  const items = iSh.getDataRange().getValues().slice(1);

  const startStr = String(startDate).trim();
  const endStr = String(endDate).trim();

  const filtered = orders.filter(r => {
    var rowDate = r[1];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    } else {
      rowDate = String(rowDate).trim();
    }
    return rowDate >= startStr && rowDate <= endStr && String(r[12]) !== 'cancelled';
  });

  const totalSales = filtered.reduce((s, r) => s + Number(r[8]), 0);
  const totalDiscount = filtered.reduce((s, r) => s + Number(r[6]), 0);
  const orderCount = filtered.length;
  const avgOrder = orderCount ? totalSales / orderCount : 0;
  const customerCount = filtered.reduce((s, r) => s + (Number(r[13]) || 1), 0);

  // Top products
  const ordIds = new Set(filtered.map(r => r[0]));
  const fItems = items.filter(r => ordIds.has(r[0]));
  const prodMap = {};
  fItems.forEach(r => {
    const k = r[2];
    if (!prodMap[k]) prodMap[k] = { name: k, qty: 0, revenue: 0 };
    prodMap[k].qty += Number(r[3]);
    prodMap[k].revenue += Number(r[5]);
  });
  const topProducts = Object.values(prodMap).sort((a, b) => b.revenue - a.revenue).slice(0, 10);

  // Payment breakdown
  const payMap = {};
  filtered.forEach(r => { payMap[r[9]] = (payMap[r[9]] || 0) + Number(r[8]); });

  // Order type breakdown
  const typeMap = {};
  filtered.forEach(r => { typeMap[r[4]] = (typeMap[r[4]] || 0) + 1; });

  // Daily breakdown
  const dailyMap = {};
  filtered.forEach(r => {
    var dk = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[1]).trim();
    if (!dailyMap[dk]) dailyMap[dk] = { sales: 0, orders: 0 };
    dailyMap[dk].sales += Number(r[8]);
    dailyMap[dk].orders++;
  });
  const daily = Object.entries(dailyMap)
    .sort((a, b) => a[0] > b[0] ? 1 : -1)
    .map(([d, v]) => ({ date: d, sales: v.sales, orders: v.orders }));

  // Hourly breakdown (for today)
  const hourlyMap = {};
  filtered.forEach(r => {
    const h = String(r[2]).split(':')[0];
    if (h) { hourlyMap[h] = (hourlyMap[h] || 0) + Number(r[8]); }
  });

  return {
    totalSales, totalDiscount, orderCount, avgOrder, customerCount,
    topProducts, payMap, typeMap, daily, hourlyMap
  };
}

function getTodayReport() {
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  return getReport(today, today);
}

// ── All Data (initial load) — with preloaded images ──────────
function getAllData() {
  var prods = [];
  var sett = {};
  try { prods = getProductsWithImages(); } catch (e) {
    try { prods = getProducts(); } catch (e2) { }
  }
  if (!prods.length) {
    try { setupSheets(); } catch (e) { }
    try { prods = getProductsWithImages(); } catch (e) {
      try { prods = getProducts(); } catch (e2) { }
    }
  }
  try { sett = getSettings(); } catch (e) { }
  // Auto-fix Orders header
  try {
    var oSh = getSS_().getSheetByName(SHEET_ORDERS);
    if (oSh) oSh.getRange(1, 1, 1, 14).setValues([['order_id', 'date', 'time', 'table_no', 'order_type', 'total', 'discount', 'discount_type', 'grand_total', 'payment', 'note', 'cashier', 'status', 'customer_count']]);
  } catch (e) { }
  return { products: prods, settings: sett, tables: [], todayReport: {} };
}

function debugTableOrders(tableNo) {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const tableStr = String(tableNo).trim();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][1];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    }
    if (String(rowDate) === today && String(data[i][3]).trim() === tableStr) {
      results.push({
        row: i + 1,
        orderId: String(data[i][0]),
        date: String(rowDate),
        table: String(data[i][3]),
        status: String(data[i][12]),
        total: Number(data[i][8]) || 0
      });
    }
  }
  return { table: tableNo, today: today, orders: results, totalOrders: results.length };
}

// Clear today's test orders and reset counter
function clearTestOrders() {
  const oSh = getSheet_(SHEET_ORDERS);
  const iSh = getSheet_(SHEET_ITEMS);
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  // Collect orderId of today's orders
  const orders = oSh.getDataRange().getValues();
  const todayIds = [];
  const rowsToDelete = [];

  for (let i = orders.length - 1; i >= 1; i--) {
    let rowDate = orders[i][1];
    if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    if (String(rowDate) === today) {
      todayIds.push(String(orders[i][0]).trim());
      rowsToDelete.push(i + 1);
    }
  }

  // Delete order rows (reverse to avoid shifting)
  rowsToDelete.forEach(r => oSh.deleteRow(r));

  // Delete matching items
  const items = iSh.getDataRange().getValues();
  const itemRowsToDelete = [];
  for (let i = items.length - 1; i >= 1; i--) {
    if (todayIds.includes(String(items[i][0]).trim())) {
      itemRowsToDelete.push(i + 1);
    }
  }
  itemRowsToDelete.forEach(r => iSh.deleteRow(r));

  // Reset counter
  const settSh = getSheet_(SHEET_SETTINGS);
  const settData = settSh.getDataRange().getValues();
  for (let i = 1; i < settData.length; i++) {
    if (String(settData[i][0]).trim() === 'last_order_num') {
      settSh.getRange(i + 1, 2).setValue(0);
      break;
    }
  }

  SpreadsheetApp.flush();
  return { cleared: rowsToDelete.length, items: itemRowsToDelete.length, date: today };
}

// Remove duplicate products - keep first occurrence of each ID
function deduplicateProducts() {
  const sh = getSheet_(SHEET_PRODUCTS);
  const data = sh.getDataRange().getValues();
  const seen = {};
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (!id) { rowsToDelete.push(i + 1); continue; }
    if (seen[id]) {
      rowsToDelete.push(i + 1); // duplicate - mark for deletion
    } else {
      seen[id] = true;
    }
  }

  // Delete in reverse to avoid row shifting
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sh.deleteRow(rowsToDelete[i]);
  }

  SpreadsheetApp.flush();
  return { removed: rowsToDelete.length, remaining: Object.keys(seen).length };
}

// ── Food Image Management (Google Drive) ─────────────────────
function setupFoodImageFolder() {
  var folders = DriveApp.getFoldersByName('POS_FoodImages');
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder('POS_FoodImages');
  }
  return { folderId: folder.getId(), folderUrl: folder.getUrl() };
}

function saveFoodImage(productId, base64Data) {
  var folders = DriveApp.getFoldersByName('POS_FoodImages');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('POS_FoodImages');
  var existing = folder.getFilesByName(productId + '.png');
  while (existing.hasNext()) existing.next().setTrashed(true);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', productId + '.png');
  var file = folder.createFile(blob);
  // Try to share publicly, skip if blocked by domain
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  var url = 'https://drive.google.com/uc?id=' + file.getId();
  updateProductImageUrl_(productId, url);
  return { success: true, url: url, fileId: file.getId() };
}

function updateProductImageUrl_(productId, url) {
  var sh = getSheet_(SHEET_PRODUCTS);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var imgCol = headers.indexOf('image_url');
  if (imgCol === -1) {
    imgCol = headers.length;
    sh.getRange(1, imgCol + 1).setValue('image_url');
  }
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(productId).trim()) {
      sh.getRange(i + 1, imgCol + 1).setValue(url);
      break;
    }
  }
}

function getAllFoodImageUrls() {
  var sh = getSheet_(SHEET_PRODUCTS);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var imgCol = headers.indexOf('image_url');
  if (imgCol === -1) return {};
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][imgCol]) map[String(data[i][0])] = String(data[i][imgCol]);
  }
  return map;
}

// Scan Drive folder and match files to products
function syncFoodImages() {
  try {
    var folders = DriveApp.getFoldersByName('POS_FoodImages');
    if (!folders.hasNext()) return { success: false, message: 'ไม่พบโฟลเดอร์ POS_FoodImages' };
    var folder = folders.next();

    // Collect files first
    var fileList = [];
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({ name: f.getName(), id: f.getId() });
    }

    var sh = getSheet_(SHEET_PRODUCTS);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var imgCol = headers.indexOf('image_url');
    if (imgCol === -1) {
      imgCol = headers.length;
      sh.getRange(1, imgCol + 1).setValue('image_url');
    }

    var prodMap = {};
    for (var i = 1; i < data.length; i++) {
      prodMap[String(data[i][0]).trim().toUpperCase()] = i + 1;
    }

    var matched = 0;
    var skipped = [];
    for (var j = 0; j < fileList.length; j++) {
      var fname = fileList[j].name.replace(/\.\w+$/, '').trim().toUpperCase();
      if (prodMap[fname]) {
        sh.getRange(prodMap[fname], imgCol + 1).setValue('drive:' + fileList[j].id);
        matched++;
      } else {
        skipped.push(fileList[j].name);
      }
    }
    SpreadsheetApp.flush();
    return { success: true, matched: matched, skipped: skipped, total: fileList.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Serve image as base64 via Drive API v3 (bypasses DriveApp restriction)
function getImageData(fileId) {
  try {
    var url = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      return 'ERR:HTTP' + response.getResponseCode();
    }
    var blob = response.getBlob();
    var bytes = blob.getBytes();
    var type = blob.getContentType() || 'image/jpeg';
    // Convert PNG > 300KB to JPEG
    if (type === 'image/png' && bytes.length > 300000) {
      try {
        blob = blob.getAs('image/jpeg');
        bytes = blob.getBytes();
        type = 'image/jpeg';
      } catch (ce) { }
    }
    if (bytes.length > 8000000) return 'ERR:file_too_large';
    return 'data:' + type + ';base64,' + Utilities.base64Encode(bytes);
  } catch (e) {
    return 'ERR:' + e.message;
  }
}

// Cached version — stores in CacheService for 6 hours
function getCachedImageData_(fileId) {
  var cache = CacheService.getScriptCache();
  var key = 'img_' + fileId;
  var cached = cache.get(key);
  if (cached) return cached;
  var data = getImageData(fileId);
  if (data && data.indexOf('data:') === 0 && data.length < 100000) {
    try { cache.put(key, data, 21600); } catch(e) {}
  }
  return data;
}

// DIAGNOSTIC: Run from Editor to test Drive image access
function testImageLoad() {
  var prods = getProducts().filter(function (p) { return p.image_url; });
  if (!prods.length) return 'ไม่มีสินค้าที่มี image_url';
  var p = prods[0];
  var fid = resolveDriveId_(p.image_url);
  Logger.log('Testing product: ' + p.id + ' | image_url: ' + p.image_url + ' | resolvedId: ' + fid);
  if (!fid) return 'Could not resolve Drive ID for: ' + p.image_url;
  var result = getImageData(fid);
  Logger.log('Result: ' + p.id + ' | dataLen: ' + result.length + ' | prefix: ' + result.substring(0, 50));
  return 'product: ' + p.id + ' | dataLen: ' + result.length + ' | prefix: ' + result.substring(0, 50);
}

// DIAGNOSTIC: Test URL resolution for ALL products that have image_url
function testResolveAllUrls() {
  var prods = getProducts().filter(function (p) { return p.image_url; });
  var results = [];
  for (var i = 0; i < prods.length; i++) {
    var p = prods[i];
    var url = String(p.image_url).trim();
    var resolved = resolveDriveId_(url);
    var status = 'NO_RESOLVE';
    if (resolved) {
      try {
        var imgData = getImageData(resolved);
        status = imgData.indexOf('data:') === 0 ? 'OK (len=' + imgData.length + ')' : imgData.substring(0, 80);
      } catch (e) {
        status = 'FETCH_ERR: ' + e.message;
      }
    }
    var line = p.id + ' | url=' + url.substring(0, 60) + ' | resolved=' + resolved + ' | ' + status;
    Logger.log(line);
    results.push(line);
  }
  return results.join('\n');
}