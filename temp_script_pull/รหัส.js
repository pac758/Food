// ============================================================
//  POS SYSTEM - Google Apps Script Backend
//  ร้านอาหาร/เครื่องดื่ม
// ============================================================

const SHEET_PRODUCTS  = 'Products';
const SHEET_ORDERS    = 'Orders';
const SHEET_ITEMS     = 'OrderItems';
const SHEET_SETTINGS  = 'Settings';

// ── Entry Point ──────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('POS System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Sheet Helpers ─────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// ── Initial Setup ─────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Products
  let p = ss.getSheetByName(SHEET_PRODUCTS);
  if (!p) {
    p = ss.insertSheet(SHEET_PRODUCTS);
    p.appendRow(['id','name','category','price','stock','unit','active']);
    p.getRange(1,1,1,7).setFontWeight('bold').setBackground('#FF6B35').setFontColor('#fff');
    // Sample data
    const sample = [
      ['P001','ข้าวผัดกุ้ง','อาหาร',80,50,'จาน',true],
      ['P002','ผัดไทย','อาหาร',70,50,'จาน',true],
      ['P003','ต้มยำกุ้ง','อาหาร',120,30,'ชาม',true],
      ['P004','น้ำเปล่า','เครื่องดื่ม',15,100,'ขวด',true],
      ['P005','น้ำส้ม','เครื่องดื่ม',35,50,'แก้ว',true],
      ['P006','ชาเย็น','เครื่องดื่ม',40,50,'แก้ว',true],
      ['P007','กาแฟเย็น','เครื่องดื่ม',45,50,'แก้ว',true],
    ];
    sample.forEach(r => p.appendRow(r));
  }

  // Orders
  let o = ss.getSheetByName(SHEET_ORDERS);
  if (!o) {
    o = ss.insertSheet(SHEET_ORDERS);
    o.appendRow(['order_id','date','time','total','discount','grand_total','payment','note','cashier']);
    o.getRange(1,1,1,9).setFontWeight('bold').setBackground('#FF6B35').setFontColor('#fff');
  }

  // OrderItems
  let i = ss.getSheetByName(SHEET_ITEMS);
  if (!i) {
    i = ss.insertSheet(SHEET_ITEMS);
    i.appendRow(['order_id','product_id','product_name','qty','unit_price','subtotal']);
    i.getRange(1,1,1,6).setFontWeight('bold').setBackground('#FF6B35').setFontColor('#fff');
  }

  // Settings
  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) {
    s = ss.insertSheet(SHEET_SETTINGS);
    s.appendRow(['key','value']);
    s.appendRow(['shop_name','ร้านอาหารของฉัน']);
    s.appendRow(['vat','0']);
    s.appendRow(['last_order_num','0']);
  }

  return { success: true, message: 'Setup complete!' };
}

// ── Settings ──────────────────────────────────────────────────
function getSettings() {
  const sh = getSheet(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const s = {};
  data.slice(1).forEach(r => { s[r[0]] = r[1]; });
  return s;
}

function saveSetting(key, value) {
  const sh = getSheet(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i+1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

// ── Products ──────────────────────────────────────────────────
function getProducts() {
  const sh = getSheet(SHEET_PRODUCTS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .filter(r => r[6] !== false)
    .map(r => ({
      id: r[0], name: r[1], category: r[2],
      price: Number(r[3]), stock: Number(r[4]),
      unit: r[5], active: r[6]
    }));
}

function getCategories() {
  const products = getProducts();
  const cats = [...new Set(products.map(p => p.category))];
  return cats;
}

function addProduct(data) {
  const sh = getSheet(SHEET_PRODUCTS);
  const id = 'P' + String(Date.now()).slice(-6);
  sh.appendRow([id, data.name, data.category, Number(data.price), Number(data.stock), data.unit, true]);
  return { success: true, id };
}

function updateProduct(data) {
  const sh = getSheet(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sh.getRange(i+1, 2, 1, 5).setValues([[data.name, data.category, Number(data.price), Number(data.stock), data.unit]]);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบสินค้า' };
}

function deleteProduct(id) {
  const sh = getSheet(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sh.getRange(i+1, 7).setValue(false);
      return { success: true };
    }
  }
  return { success: false };
}

function updateStock(id, delta) {
  const sh = getSheet(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      const newStock = Math.max(0, Number(rows[i][4]) + delta);
      sh.getRange(i+1, 5).setValue(newStock);
      return newStock;
    }
  }
}

// ── Orders ────────────────────────────────────────────────────
function generateOrderId() {
  const sh = getSheet(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  let num = 0;
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'last_order_num') {
      num = Number(data[i][1]) + 1;
      row = i + 1;
      break;
    }
  }
  if (row > 0) sh.getRange(row, 2).setValue(num);
  const d = new Date();
  const prefix = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyyMMdd');
  return `${prefix}-${String(num).padStart(4,'0')}`;
}

function saveOrder(orderData) {
  try {
    const orderId = generateOrderId();
    const now = new Date();
    const date = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
    const time = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

    const oSh = getSheet(SHEET_ORDERS);
    oSh.appendRow([
      orderId, date, time,
      orderData.total, orderData.discount || 0,
      orderData.grandTotal, orderData.payment,
      orderData.note || '', orderData.cashier || 'แคชเชียร์'
    ]);

    const iSh = getSheet(SHEET_ITEMS);
    orderData.items.forEach(item => {
      iSh.appendRow([orderId, item.id, item.name, item.qty, item.price, item.qty * item.price]);
      updateStock(item.id, -item.qty);
    });

    const settings = getSettings();
    return { success: true, orderId, shopName: settings.shop_name || 'ร้านของฉัน', date, time };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ── Reports ───────────────────────────────────────────────────
function getReport(type, startDate, endDate) {
  const oSh = getSheet(SHEET_ORDERS);
  const iSh = getSheet(SHEET_ITEMS);

  const orders = oSh.getDataRange().getValues().slice(1);
  const items  = iSh.getDataRange().getValues().slice(1);

  const start = new Date(startDate); start.setHours(0,0,0,0);
  const end   = new Date(endDate);   end.setHours(23,59,59,999);

  const filtered = orders.filter(r => {
    const d = new Date(r[1]);
    return d >= start && d <= end;
  });

  const totalSales  = filtered.reduce((s, r) => s + Number(r[5]), 0);
  const orderCount  = filtered.length;
  const avgOrder    = orderCount ? totalSales / orderCount : 0;

  // Best sellers
  const ordIds = new Set(filtered.map(r => r[0]));
  const filteredItems = items.filter(r => ordIds.has(r[0]));
  const productMap = {};
  filteredItems.forEach(r => {
    const k = r[2];
    if (!productMap[k]) productMap[k] = { name: k, qty: 0, revenue: 0 };
    productMap[k].qty     += Number(r[3]);
    productMap[k].revenue += Number(r[5]);
  });
  const topProducts = Object.values(productMap).sort((a,b) => b.revenue - a.revenue).slice(0,10);

  // Payment breakdown
  const payMap = {};
  filtered.forEach(r => {
    payMap[r[6]] = (payMap[r[6]] || 0) + Number(r[5]);
  });

  // Daily breakdown (for charts)
  const dailyMap = {};
  filtered.forEach(r => {
    const d = r[1];
    if (!dailyMap[d]) dailyMap[d] = 0;
    dailyMap[d] += Number(r[5]);
  });
  const daily = Object.entries(dailyMap).sort((a,b) => a[0] > b[0] ? 1 : -1);

  return { totalSales, orderCount, avgOrder, topProducts, payMap, daily };
}

function getTodayReport() {
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  return getReport('today', today, today);
}