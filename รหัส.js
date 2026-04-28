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

  // ─── Products (ล้างแล้วใส่ใหม่ทุกครั้ง) ───
  let p = ss.getSheetByName(SHEET_PRODUCTS);
  if (!p) p = ss.insertSheet(SHEET_PRODUCTS);
  {
    p.clear();
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

  // ─── OrderItems ───
  let i = ss.getSheetByName(SHEET_ITEMS);
  if (!i) i = ss.insertSheet(SHEET_ITEMS);
  if (!i.getRange('A1').getValue()) {
    i.appendRow(['order_id','product_id','product_name','qty',
                 'unit_price','subtotal','spice_level','options','note']);
    headerStyle(i, 9);
  }

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
      sort_order: Number(r[11]) || 99
    }))
    .sort((a,b) => a.sort_order - b.sort_order);
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
    if (data[i][0] === 'last_order_num') { num = Number(data[i][1]) + 1; row = i+1; break; }
  }
  if (row > 0) sh.getRange(row, 2).setValue(num);
  const d = new Date();
  const prefix = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyyMMdd');
  return `${prefix}-${String(num).padStart(4,'0')}`;
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
        item.spice || '', item.options || '', item.note || ''
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

function updateOrderStatus(orderId, status) {
  const sh = getSheet_(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  const oid = String(orderId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === oid) {
      var oldStatus = String(data[i][12]);
      sh.getRange(i+1, 13).setValue(status);
      SpreadsheetApp.flush();
      return { success: true, oldStatus: oldStatus, newStatus: status, row: i+1 };
    }
  }
  return { success: false, message: 'ไม่พบ Order: ' + oid + ' (rows: ' + data.length + ')' };
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
    .filter(r => r[0] === orderId)
    .map(r => ({
      product_id: r[1], product_name: r[2], qty: Number(r[3]),
      unit_price: Number(r[4]), subtotal: Number(r[5]),
      spice: r[6], options: r[7], note: r[8]
    }));
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
      const activeOrder = orders.slice(1).find(r => {
        let rowDate = r[1];
        if (rowDate instanceof Date) {
          rowDate = Utilities.formatDate(rowDate, 'Asia/Bangkok', 'yyyy-MM-dd');
        }
        return String(rowDate) === today && String(r[3]).trim() === tableStr && ['new','cooking','served'].includes(String(r[12]));
      });

      if (activeOrder) {
        var orderTime = activeOrder[2];
        if (orderTime instanceof Date) {
          orderTime = Utilities.formatDate(orderTime, 'Asia/Bangkok', 'HH:mm');
        } else {
          orderTime = String(orderTime || '');
        }
        tables.push({
          no: t,
          status: String(activeOrder[12]),
          orderId: String(activeOrder[0]),
          orderType: String(activeOrder[4]),
          total: Number(activeOrder[8]) || 0,
          time: orderTime,
          itemCount: countOrderItems_(String(activeOrder[0]))
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
  return { products: prods, settings: sett, tables: [], todayReport: {}, debug: debug.join(' | ') };
}