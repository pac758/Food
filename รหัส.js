// ============================================================
//  POS SYSTEM — ร้านลาบบ้านสวน
//  Professional Restaurant POS Backend
//  Google Apps Script + Google Sheets
// ============================================================

const SHEET_PRODUCTS = 'Products';
const SHEET_ORDERS   = 'Orders';
const SHEET_ITEMS    = 'OrderItems';
const SHEET_SETTINGS = 'Settings';
const SHEET_RESERVATIONS = 'Reservations';
const SHEET_STAFF    = 'Staff';

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

// ── Google Sheet ID (fallback) ────────────────────────────────
const SPREADSHEET_ID = '1V0j6xcnjrQ2XWy3MGHYM5NbX9yS14iHJn6Njm_LS1OY';

// ── Sheet Helpers ─────────────────────────────────────────────
function getSS_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
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
    sh.getRange(1,1,1,cols).setFontWeight('bold')
      .setBackground('#FF6B35').setFontColor('#fff');
    sh.setFrozenRows(1);
  };

  // ─── Products (สร้างครั้งแรกเท่านั้น ไม่ล้างของเดิม) ───
  let p = ss.getSheetByName(SHEET_PRODUCTS);
  if (!p) p = ss.insertSheet(SHEET_PRODUCTS);
  if (!p.getRange('A1').getValue()) {
    p.appendRow(['id','name','category','price','stock','unit','active',
                 'emoji','spice_default','options','cost','sort_order']);
    headerStyle(p, 12);

    const menu = [
      // ─── 🌟 เมนูแนะนำ ───
      ['M001','ลาบหมูคั่ว','🌟 แนะนำ',80,99,'จาน',true,'🥩','กลาง','เพิ่มไข่มดแดง:20',35,1],
      ['M002','ต้มแซ่บกระดูกอ่อน','🌟 แนะนำ',90,99,'ชาม',true,'🍲','กลาง','',40,2],
      ['M003','คอหมูย่าง','🌟 แนะนำ',100,99,'จาน',true,'🔥','','' ,45,3],

      // ─── ลาบ / น้ำตก ───
      ['M004','ลาบหมู','ลาบ/น้ำตก',70,99,'จาน',true,'🥩','กลาง','เพิ่มไข่มดแดง:20',30,10],
      ['M005','ลาบไก่','ลาบ/น้ำตก',70,99,'จาน',true,'🍗','กลาง','เพิ่มไข่มดแดง:20',30,11],
      ['M006','ลาบเนื้อ','ลาบ/น้ำตก',90,99,'จาน',true,'🥩','กลาง','เพิ่มไข่มดแดง:20',40,12],
      ['M007','น้ำตกหมู','ลาบ/น้ำตก',80,99,'จาน',true,'🥩','กลาง','',35,13],
      ['M008','น้ำตกเนื้อ','ลาบ/น้ำตก',100,99,'จาน',true,'🥩','กลาง','',45,14],

      // ─── ต้มแซ่บ ───
      ['M009','ต้มแซ่บไก่บ้าน','ต้มแซ่บ',100,99,'ชาม',true,'🍲','กลาง','',45,20],
      ['M010','ต้มแซ่บเนื้อ','ต้มแซ่บ',120,99,'ชาม',true,'🍲','กลาง','',55,21],
      ['M011','อ่อมหมู','ต้มแซ่บ',90,99,'ชาม',true,'🍲','กลาง','',40,22],
      ['M012','อ่อมเนื้อ','ต้มแซ่บ',90,99,'ชาม',true,'🍲','กลาง','',40,23],

      // ─── เมนูย่าง ───
      ['M013','ไก่ย่างบ้านสวน','เมนูย่าง',120,99,'ตัว',true,'🍗','','',50,30],
      ['M014','เสือร้องไห้','เมนูย่าง',120,99,'จาน',true,'🥩','','',55,31],

      // ─── เครื่องเคียง ───
      ['M015','ข้าวเหนียว','เครื่องเคียง',10,999,'ห่อ',true,'🍚','','',3,40],
      ['M016','ผักสด/ผักลวก','เครื่องเคียง',0,999,'จาน',true,'🥬','','',5,41],
      ['M017','แจ่วปลาร้า','เครื่องเคียง',15,999,'ถ้วย',true,'🫙','','',5,42],
      ['M018','แจ่วบอง','เครื่องเคียง',15,999,'ถ้วย',true,'🫙','','',5,43],

      // ─── เครื่องดื่ม ───
      ['M019','น้ำเปล่า','เครื่องดื่ม',15,200,'ขวด',true,'💧','','',7,50],
      ['M020','น้ำอัดลม','เครื่องดื่ม',20,200,'กระป๋อง',true,'🥤','','',10,51],
      ['M021','ชาเย็น','เครื่องดื่ม',40,99,'แก้ว',true,'🧋','','',15,52],
      ['M022','กาแฟเย็น','เครื่องดื่ม',45,99,'แก้ว',true,'☕','','',18,53],
      ['M023','น้ำมะนาว','เครื่องดื่ม',35,99,'แก้ว',true,'🍋','','',12,54],
      ['M024','โอเลี้ยง','เครื่องดื่ม',35,99,'แก้ว',true,'☕','','',12,55],
    ];
    menu.forEach(r => p.appendRow(r));
  }

  // ─── Orders ───
  let o = ss.getSheetByName(SHEET_ORDERS);
  if (!o) o = ss.insertSheet(SHEET_ORDERS);
  if (!o.getRange('A1').getValue()) {
    o.appendRow(['order_id','date','time','table_no','order_type',
                 'total','discount','discount_type','grand_total',
                 'payment','note','cashier','status','customer_count']);
    headerStyle(o, 14);
  }
  // Always fix header row to correct format
  const correctHeaders = ['order_id','date','time','table_no','order_type',
    'total','discount','discount_type','grand_total',
    'payment','note','cashier','status','customer_count'];
  o.getRange(1, 1, 1, 14).setValues([correctHeaders]);

  // ─── OrderItems ───
  let i = ss.getSheetByName(SHEET_ITEMS);
  if (!i) i = ss.insertSheet(SHEET_ITEMS);
  if (!i.getRange('A1').getValue()) {
    i.appendRow(['order_id','product_id','product_name','qty',
                 'unit_price','subtotal','spice_level','options','note','item_status']);
    headerStyle(i, 10);
  }
  // Auto-fix OrderItems header
  i.getRange(1,1,1,10).setValues([['order_id','product_id','product_name','qty','unit_price','subtotal','spice_level','options','note','item_status']]);

  // ─── Settings ───
  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) s = ss.insertSheet(SHEET_SETTINGS);
  if (!s.getRange('A1').getValue()) {
    s.appendRow(['key','value']);
    headerStyle(s, 2);
    const defaults = [
      ['shop_name','ลาบบ้านสวน'],
      ['shop_phone','090-123-4567'],
      ['vat','0'],
      ['last_order_num','0'],
      ['table_count','10'],
      ['promptpay_id','0901234567'],
      ['promo_family_price','299'],
      ['promo_student_discount','10'],
      ['promo_checkin_enabled','true'],
    ];
    defaults.forEach(r => s.appendRow(r));
  }

  // ─── Reservations ───
  let r = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!r) r = ss.insertSheet(SHEET_RESERVATIONS);
  if (!r.getRange('A1').getValue()) {
    r.appendRow(['reservation_id','date','time','table_no','customer_name','party_size','contact','note','status','created_by','created_at']);
    headerStyle(r, 11);
  }

  // ─── Staff ───
  let st = ss.getSheetByName(SHEET_STAFF);
  if (!st) st = ss.insertSheet(SHEET_STAFF);
  if (!st.getRange('A1').getValue()) {
    st.appendRow(['name','pin','role','active']);
    headerStyle(st, 4);
    st.appendRow(['พนักงาน 1','1234','cashier',true]);
    st.appendRow(['พนักงาน 2','5678','cashier',true]);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Setup สำเร็จ!' };
}

// ── Staff / Login ────────────────────────────────────────────
function loginWithPin(pin) {
  let staff = sheetToObjects_(SHEET_STAFF, ['name','pin','role','active']);
  // Auto-setup if no staff exists
  if (!staff.length) {
    setupSheets();
    staff = sheetToObjects_(SHEET_STAFF, ['name','pin','role','active']);
  }
  const found = staff.find(s => String(s.pin) === String(pin) && s.active !== false);
  if (found) return { success: true, name: found.name, role: found.role };
  return { success: false };
}

function getStaffList() {
  return sheetToObjects_(SHEET_STAFF, ['name','pin','role','active'])
    .filter(s => s.active !== false)
    .map(s => ({ name: s.name, role: s.role }));
}

// ── Settings ─────────────────────────────────────────────────
function getSettings() {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const s = {};
  data.slice(1).forEach(r => { s[r[0]] = r[1]; });
  return s;
}

function saveSetting(key, value) {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i+1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

function saveMultipleSettings(kvPairs) {
  kvPairs.forEach(([k,v]) => saveSetting(k, v));
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
    .sort((a,b) => a.sort_order - b.sort_order);
}

// Customer menu - preload images server-side
function getMenuForCustomer() {
  const prods = getProducts().map(p => {
    var imgData = '';
    if (p.image_url && String(p.image_url).indexOf('drive:') === 0) {
      try {
        var fid = String(p.image_url).replace('drive:', '');
        imgData = getImageData(fid);
      } catch(e) { imgData = 'ERR:' + e.message; }
    }
    return {
      id: p.id, name: p.name, category: p.category,
      price: p.price, unit: p.unit, emoji: p.emoji,
      spice_default: p.spice_default, options: p.options,
      image_data: imgData,
      has_drive: p.image_url ? String(p.image_url).substring(0,40) : ''
    };
  });
  const sett = getSettings();
  return { products: prods, shopName: sett.shop_name || 'ลาบบ้านสวน' };
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
      sh.getRange(i+1, 2, 1, 11).setValues([[
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
    if (rows[i][0] === id) { sh.getRange(i+1, 7).setValue(false); return { success: true }; }
  }
  return { success: false };
}

function updateStock(id, delta) {
  const sh = getSheet_(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      const ns = Math.max(0, Number(rows[i][4]) + delta);
      sh.getRange(i+1, 5).setValue(ns);
      return ns;
    }
  }
}

// ── Orders ───────────────────────────────────────────────────
function generateOrderId_() {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  let num = 0, row = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'last_order_num') { num = Number(data[i][1]) || 0; row = i+1; break; }
  }
  num++;
  if (row > 0) {
    sh.getRange(row, 2).setValue(num);
  } else {
    // Create the setting if it doesn't exist
    sh.appendRow(['last_order_num', num]);
  }
  SpreadsheetApp.flush();
  const d = new Date();
  const prefix = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyyMMdd');
  return prefix + '-' + String(num).padStart(4,'0');
}

function saveOrder(orderData) {
  try {
    const orderId = generateOrderId_();
    const now = new Date();
    const date = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
    const time = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

    const oSh = getSheet_(SHEET_ORDERS);
    oSh.appendRow([
      orderId, date, time,
      orderData.tableNo || '-',
      orderData.orderType || 'dine-in',
      orderData.total,
      orderData.discount || 0,
      orderData.discountType || '',
      orderData.grandTotal,
      orderData.payment || 'เงินสด',
      orderData.note || '',
      orderData.cashier || '',
      'new',
      orderData.customerCount || 1
    ]);

    const iSh = getSheet_(SHEET_ITEMS);
    orderData.items.forEach(item => {
      iSh.appendRow([
        orderId, item.id, item.name, item.qty,
        item.price, item.qty * item.price,
        item.spice || '', item.options || '', item.note || '', 'cooking'
      ]);
      updateStock(item.id, -item.qty);
    });

    const settings = getSettings();
    return {
      success: true, orderId,
      shopName: settings.shop_name || 'ลาบบ้านสวน',
      shopPhone: settings.shop_phone || '',
      date, time
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function updateOrderStatus(orderId, status, tableNo) {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  const oid = String(orderId).trim();
  const tbl = tableNo ? String(tableNo).trim() : '';
  // Reverse: find LATEST matching row
  for (let i = data.length - 1; i >= 1; i--) {
    var matchId = String(data[i][0]).trim() === oid;
    var matchTbl = !tbl || String(data[i][3]).trim().replace(/['"]/g,'') === tbl;
    if (matchId && matchTbl) {
      var oldStatus = String(data[i][12]);
      sh.getRange(i+1, 13).setValue(status);
      SpreadsheetApp.flush();
      return { success: true, oldStatus: oldStatus, newStatus: status, row: i+1, table: String(data[i][3]) };
    }
  }
  return { success: false, message: 'ไม่พบ Order: ' + oid + ' table: ' + tbl };
}

function cancelOrder(orderId, reason) {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  const oid = String(orderId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === oid) {
      sh.getRange(i+1, 13).setValue('cancelled');
      const note = data[i][10] ? data[i][10] + ' | ยกเลิก: ' + reason : 'ยกเลิก: ' + reason;
      sh.getRange(i+1, 11).setValue(note);
      // Restore stock
      const items = getOrderItems_(orderId);
      items.forEach(item => updateStock(item.product_id, item.qty));
      return { success: true };
    }
  }
  return { success: false };
}

function getOrderItems_(orderId) {
  const sh = getSheet_(SHEET_ITEMS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .map((r, idx) => ({ r, rowIdx: idx + 2 }))
    .filter(x => x.r[0] === orderId)
    .map(x => ({
      product_id: x.r[1], product_name: x.r[2], qty: Number(x.r[3]),
      unit_price: Number(x.r[4]), subtotal: Number(x.r[5]),
      spice: x.r[6], options: x.r[7], note: x.r[8],
      item_status: x.r[9] || 'cooking',
      row: x.rowIdx
    }));
}

// Update single item status
function updateItemStatus(orderId, itemRow, newStatus) {
  const iSh = getSheet_(SHEET_ITEMS);
  iSh.getRange(itemRow, 10).setValue(newStatus);
  
  // Check if ALL items of this order are served
  const data = iSh.getDataRange().getValues();
  const orderItems = data.slice(1).filter(r => String(r[0]) === String(orderId));
  const allServed = orderItems.length > 0 && orderItems.every(r => String(r[9]) === 'served');
  
  // Auto-update order status
  if (allServed) {
    const oSh = getSheet_(SHEET_ORDERS);
    const oData = oSh.getDataRange().getValues();
    for (let i = oData.length - 1; i >= 1; i--) {
      if (String(oData[i][0]).trim() === String(orderId).trim()) {
        oSh.getRange(i+1, 13).setValue('served');
        break;
      }
    }
  }
  
  SpreadsheetApp.flush();
  return { success: true, allServed: allServed };
}

// ── Kitchen Display ──────────────────────────────────────────
function getKitchenOrders() {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const activeStatuses = ['new','cooking'];

  const orders = data.slice(1)
    .filter(r => r[1] === today && activeStatuses.includes(r[12]))
    .map(r => ({
      orderId: r[0], date: r[1], time: r[2],
      tableNo: r[3], orderType: r[4],
      status: r[12],
      items: getOrderItems_(r[0])
    }));

  return orders.sort((a,b) => a.time > b.time ? 1 : -1);
}

// ── Order History ────────────────────────────────────────────
function getOrderHistory(days) {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - (days || 7));
  cutoff.setHours(0,0,0,0);

  return data.slice(1)
    .filter(r => new Date(r[1]) >= cutoff)
    .map(r => ({
      orderId: r[0], date: r[1], time: r[2],
      tableNo: r[3], orderType: r[4],
      total: Number(r[5]), discount: Number(r[6]),
      discountType: r[7], grandTotal: Number(r[8]),
      payment: r[9], note: r[10], cashier: r[11],
      status: r[12], customerCount: Number(r[13])
    }))
    .reverse();
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
        return String(rowDate) === today && String(r[3]).trim().replace(/['"]/g,'') === tableStr && ['new','cooking','served'].includes(String(r[12]));
      });
      // Priority: cooking/served orders first (active), then new
      const activeOrder = todayOrders.filter(r => ['cooking','served'].includes(String(r[12]))).pop()
                       || todayOrders.filter(r => String(r[12]) === 'new').pop()
                       || null;

      if (activeOrder) {
        var orderTime = activeOrder[2];
        if (orderTime instanceof Date) {
          orderTime = Utilities.formatDate(orderTime, 'Asia/Bangkok', 'HH:mm');
        } else {
          orderTime = String(orderTime || '');
        }
        var oid = String(activeOrder[0]);
        var itemInfo = getItemCounts_(oid);
        tables.push({
          no: t,
          status: String(activeOrder[12]),
          orderId: oid,
          orderType: String(activeOrder[4]),
          total: Number(activeOrder[8]) || 0,
          time: orderTime,
          itemCount: itemInfo.total,
          servedCount: itemInfo.served,
          totalItems: itemInfo.total
        });
        continue;
      }

      const reservation = reservations.find(r => String(r.tableNo).trim() === tableStr && ['reserved','arrived','seated'].includes(r.status));
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
          itemCount: 0
        });
        continue;
      }

      tables.push({
        no: t,
        status: 'available',
        orderId: '',
        total: 0,
        time: '',
        itemCount: 0
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

function countOrderItems_(orderId) {
  try {
    const sh = getSheet_(SHEET_ITEMS);
    const data = sh.getDataRange().getValues();
    return data.slice(1).filter(r => String(r[0]) === orderId).length;
  } catch(e) { return 0; }
}

// ── Reports ──────────────────────────────────────────────────
function getReport(startDate, endDate) {
  const oSh = getSheet_(SHEET_ORDERS);
  const iSh = getSheet_(SHEET_ITEMS);
  const orders = oSh.getDataRange().getValues().slice(1);
  const items  = iSh.getDataRange().getValues().slice(1);

  const start = new Date(startDate); start.setHours(0,0,0,0);
  const end   = new Date(endDate);   end.setHours(23,59,59,999);

  const filtered = orders.filter(r => {
    const d = new Date(r[1]);
    return d >= start && d <= end && r[12] !== 'cancelled';
  });

  const totalSales = filtered.reduce((s,r) => s + Number(r[8]), 0);
  const totalDiscount = filtered.reduce((s,r) => s + Number(r[6]), 0);
  const orderCount = filtered.length;
  const avgOrder = orderCount ? totalSales / orderCount : 0;
  const customerCount = filtered.reduce((s,r) => s + (Number(r[13]) || 1), 0);

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
  const topProducts = Object.values(prodMap).sort((a,b) => b.revenue - a.revenue).slice(0,10);

  // Payment breakdown
  const payMap = {};
  filtered.forEach(r => { payMap[r[9]] = (payMap[r[9]] || 0) + Number(r[8]); });

  // Order type breakdown
  const typeMap = {};
  filtered.forEach(r => { typeMap[r[4]] = (typeMap[r[4]] || 0) + 1; });

  // Daily breakdown
  const dailyMap = {};
  filtered.forEach(r => {
    if (!dailyMap[r[1]]) dailyMap[r[1]] = { sales: 0, orders: 0 };
    dailyMap[r[1]].sales += Number(r[8]);
    dailyMap[r[1]].orders++;
  });
  const daily = Object.entries(dailyMap)
    .sort((a,b) => a[0] > b[0] ? 1 : -1)
    .map(([d,v]) => ({ date: d, sales: v.sales, orders: v.orders }));

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

// ── All Data (initial load) ──────────────────────────────────
function getAllData() {
  var debug = [];
  var prods = [];
  var sett = {};
  try {
    var ss = getSS_();
    debug.push('SS=' + ss.getName());
    var sh = ss.getSheetByName('Products');
    debug.push('Sheet=' + (sh ? 'found,rows=' + sh.getLastRow() : 'NOT FOUND'));
    if (sh && sh.getLastRow() > 1) {
      var d = sh.getDataRange().getValues();
      debug.push('data=' + d.length + ',hdr=' + JSON.stringify(d[0]).substring(0,80));
    }
  } catch(ex) { debug.push('dbg err=' + ex); }
  try { prods = getProducts(); debug.push('prods=' + prods.length); } catch(e) { debug.push('getProd err=' + e); }
  if (!prods.length) {
    try { setupSheets(); debug.push('setup done'); } catch(e) { debug.push('setup err=' + e); }
   try { prods = getProducts(); debug.push('retry prods=' + prods.length); } catch(e) { debug.push('retry err=' + e); }
  }
  try { sett = getSettings(); } catch(e) { debug.push('sett err=' + e); }
  // Auto-fix Orders header
  try {
    var oSh = getSS_().getSheetByName(SHEET_ORDERS);
    if (oSh) oSh.getRange(1,1,1,14).setValues([['order_id','date','time','table_no','order_type','total','discount','discount_type','grand_total','payment','note','cashier','status','customer_count']]);
  } catch(e) {}
  return { products: prods, settings: sett, tables: [], todayReport: {}, debug: debug.join(' | ') };
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
      settSh.getRange(i+1, 2).setValue(0);
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
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
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
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Serve image as base64 (simplified - no cache)
function getImageData(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    var type = blob.getContentType();
    // Convert PNG > 300KB to JPEG
    if (type === 'image/png' && bytes.length > 300000) {
      try {
        blob = blob.getAs('image/jpeg');
        bytes = blob.getBytes();
        type = 'image/jpeg';
      } catch(ce) {}
    }
    if (bytes.length > 8000000) return 'ERR:file_too_large_' + bytes.length;
    return 'data:' + type + ';base64,' + Utilities.base64Encode(bytes);
  } catch(e) {
    return 'ERR:' + e.message;
  }
}

// DIAGNOSTIC: Run from Editor to test Drive image access
function testImageLoad() {
  var prods = getProducts().filter(function(p){ return p.image_url; });
  if (!prods.length) return 'ไม่มีสินค้าที่มี image_url';
  var p = prods[0];
  var fid = String(p.image_url).replace('drive:', '');
  Logger.log('Testing product: ' + p.id + ' | image_url: ' + p.image_url + ' | fileId: ' + fid);
  var result = getImageData(fid);
  Logger.log('Result: ' + p.id + ' | dataLen: ' + result.length + ' | prefix: ' + result.substring(0, 50));
  return 'product: ' + p.id + ' | dataLen: ' + result.length + ' | prefix: ' + result.substring(0, 50);
}