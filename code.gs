/**
 * SHOP BACKEND — Google Apps Script Web App
 * -----------------------------------------
 * Paste this whole file into Extensions > Apps Script (bound to your Google Sheet).
 * Then run `setup` once from the Apps Script editor (see SETUP.md), then deploy
 * as a Web App ("Execute as: Me", "Who has access: Anyone").
 *
 * Sheets used (created automatically by setup()):
 *   Products : ID | Name | Unit | CurrentStock | BuyPrice | SellPriceDiscount | SellPriceNoDiscount | LastUpdated
 *   StockIn  : Timestamp | ProductID | ProductName | Unit | Qty | BuyPrice | SellPriceDiscount | SellPriceNoDiscount | TotalDiscount | TotalNoDiscount | Staff
 *   Sales    : SaleID | Timestamp | ProductID | ProductName | Unit | Qty | CustomerType | UnitPriceUsed | BuyPriceAtSale | SellPriceNoDiscountAtSale | LineTotal | Staff
 *
 * Design note: BuyPriceAtSale and SellPriceNoDiscountAtSale are SNAPSHOTS taken
 * at sale time. Without them, editing a product's price later would silently
 * rewrite the profit/discount history of every past sale.
 */

var SHEETS = {
  PRODUCTS: 'Products',
  STOCK_IN: 'StockIn',
  SALES: 'Sales'
};

var LOW_STOCK_THRESHOLD = 5;

var DEFAULT_PRODUCTS = [
  ['Yogurt Cups', 'pcs'],
  ['Fino 1 Box', 'box'],
  ['Fino 1 PC', 'pcs'],
  ['Mtindi MD', 'pcs'],
  ['Mtindi KOPO', 'pcs'],
  ['Mtindi MK', 'pcs'],
  ['Drinks', 'pcs'],
  ['Juice Box', 'box'],
  ['Juice 1 PC', 'pcs'],
  ['Mtindi 1L', 'bottle'],
  ['Mtindi 3L', 'bottle'],
  ['Mtindi 5L', 'bottle'],
  ['Apple', 'pcs'],
  ['TBA 1 PC', 'pcs'],
  ['Bucket', 'pcs']
];

/** Run this ONCE from the Apps Script editor to create/reset the sheets. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var products = ss.getSheetByName(SHEETS.PRODUCTS) || ss.insertSheet(SHEETS.PRODUCTS);
  products.clear();
  products.appendRow(['ID', 'Name', 'Unit', 'CurrentStock', 'BuyPrice', 'SellPriceDiscount', 'SellPriceNoDiscount', 'LastUpdated']);
  DEFAULT_PRODUCTS.forEach(function (p, i) {
    products.appendRow(['P' + (i + 1), p[0], p[1], 0, 0, 0, 0, new Date()]);
  });
  products.setFrozenRows(1);

  var stockIn = ss.getSheetByName(SHEETS.STOCK_IN) || ss.insertSheet(SHEETS.STOCK_IN);
  stockIn.clear();
  stockIn.appendRow(['Timestamp', 'ProductID', 'ProductName', 'Unit', 'Qty', 'BuyPrice', 'SellPriceDiscount', 'SellPriceNoDiscount', 'TotalDiscount', 'TotalNoDiscount', 'Staff']);
  stockIn.setFrozenRows(1);

  var sales = ss.getSheetByName(SHEETS.SALES) || ss.insertSheet(SHEETS.SALES);
  sales.clear();
  sales.appendRow(['SaleID', 'Timestamp', 'ProductID', 'ProductName', 'Unit', 'Qty', 'CustomerType', 'UnitPriceUsed', 'BuyPriceAtSale', 'SellPriceNoDiscountAtSale', 'LineTotal', 'Staff']);
  sales.setFrozenRows(1);

  Logger.log('Setup complete.');
}

/* ---------------- HTTP entry points ---------------- */

function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) || 'ping';
    var result;
    if (action === 'products') {
      result = { ok: true, products: getProductsList() };
    } else if (action === 'dashboard') {
      result = { ok: true, data: getDashboard() };
    } else if (action === 'reports') {
      result = { ok: true, data: getReports(e.parameter.period || 'day') };
    } else {
      result = { ok: true, message: 'Shop API is running' };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var result;
    if (payload.action === 'registerStock') {
      result = registerStock(payload);
    } else if (payload.action === 'recordSale') {
      result = recordSale(payload);
    } else {
      result = { ok: false, error: 'Unknown action: ' + payload.action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Sheet helpers ---------------- */

function readSheetObjects(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = { __row: i + 1 };
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    rows.push(obj);
  }
  return { sheet: sheet, headers: headers, rows: rows };
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function getProductsList() {
  return readSheetObjects(SHEETS.PRODUCTS).rows.map(function (r) {
    return {
      id: r.ID,
      name: r.Name,
      unit: r.Unit,
      stock: Number(r.CurrentStock) || 0,
      buyPrice: Number(r.BuyPrice) || 0,
      sellDiscount: Number(r.SellPriceDiscount) || 0,
      sellNoDiscount: Number(r.SellPriceNoDiscount) || 0
    };
  });
}

/* ---------------- Registration (stock-in) ---------------- */

function registerStock(payload) {
  var qty = Number(payload.qty);
  var buyPrice = Number(payload.buyPrice);
  var sellDiscount = Number(payload.sellDiscount);
  var sellNoDiscount = Number(payload.sellNoDiscount);
  if (!qty || qty <= 0) return { ok: false, error: 'Quantity must be greater than zero.' };
  if (isNaN(buyPrice) || isNaN(sellDiscount) || isNaN(sellNoDiscount)) {
    return { ok: false, error: 'Prices must be numbers.' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheetObjs = readSheetObjects(SHEETS.PRODUCTS);
    var row = null;
    if (payload.productId) {
      row = sheetObjs.rows.filter(function (r) { return r.ID === payload.productId; })[0];
    }
    if (!row && payload.name) {
      var nameLower = String(payload.name).trim().toLowerCase();
      row = sheetObjs.rows.filter(function (r) { return String(r.Name).trim().toLowerCase() === nameLower; })[0];
    }

    var productId, productName, unit;
    var sheet = sheetObjs.sheet;

    if (row) {
      productId = row.ID;
      productName = row.Name;
      unit = payload.unit || row.Unit;
      var newStock = (Number(row.CurrentStock) || 0) + qty;
      sheet.getRange(row.__row, 4, 1, 5).setValues([[newStock, buyPrice, sellDiscount, sellNoDiscount, new Date()]]);
    } else {
      // brand-new product not in the master list
      var maxNum = 0;
      sheetObjs.rows.forEach(function (r) {
        var m = /^P(\d+)$/.exec(r.ID);
        if (m) maxNum = Math.max(maxNum, Number(m[1]));
      });
      productId = 'P' + (maxNum + 1);
      productName = payload.name;
      unit = payload.unit || 'pcs';
      sheet.appendRow([productId, productName, unit, qty, buyPrice, sellDiscount, sellNoDiscount, new Date()]);
    }

    var totalDiscount = round2(qty * sellDiscount);
    var totalNoDiscount = round2(qty * sellNoDiscount);
    var stockInSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.STOCK_IN);
    stockInSheet.appendRow([
      new Date(), productId, productName, unit, qty, buyPrice, sellDiscount, sellNoDiscount,
      totalDiscount, totalNoDiscount, payload.staff || ''
    ]);

    return {
      ok: true,
      product: { id: productId, name: productName, unit: unit },
      totalDiscount: totalDiscount,
      totalNoDiscount: totalNoDiscount
    };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Sales (with locked stock deduction) ---------------- */

function recordSale(payload) {
  var items = payload.items || [];
  if (!items.length) return { ok: false, error: 'Cart is empty.' };

  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000);
  if (!gotLock) return { ok: false, error: 'System is busy processing another sale — try again in a moment.' };

  try {
    var sheetObjs = readSheetObjects(SHEETS.PRODUCTS);
    var byId = {};
    sheetObjs.rows.forEach(function (r) { byId[r.ID] = r; });

    // 1) validate the whole cart BEFORE writing anything
    var errors = [];
    items.forEach(function (item) {
      var p = byId[item.productId];
      if (!p) { errors.push('Unknown product: ' + item.productId); return; }
      var qty = Number(item.qty) || 0;
      if (qty <= 0) { errors.push(p.Name + ': quantity must be greater than zero.'); return; }
      if (qty > Number(p.CurrentStock)) {
        errors.push(p.Name + ': only ' + p.CurrentStock + ' in stock, tried to sell ' + qty + '.');
      }
    });
    if (errors.length) return { ok: false, error: errors.join(' ') };

    // 2) all good — apply deductions in memory, then write once
    var saleId = 'S' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    var now = new Date();
    var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SALES);
    var lines = [];
    var total = 0;

    items.forEach(function (item) {
      var p = byId[item.productId];
      var qty = Number(item.qty);
      var unitPrice;
      if (item.manualPrice !== undefined && item.manualPrice !== null && item.manualPrice !== '') {
        unitPrice = Number(item.manualPrice);
      } else {
        unitPrice = item.customerType === 'wholesale' ? Number(p.SellPriceDiscount) : Number(p.SellPriceNoDiscount);
      }
      var lineTotal = round2(unitPrice * qty);
      total += lineTotal;

      salesSheet.appendRow([
        saleId, now, p.ID, p.Name, p.Unit, qty, item.customerType, unitPrice,
        Number(p.BuyPrice), Number(p.SellPriceNoDiscount), lineTotal, payload.staff || ''
      ]);

      p.CurrentStock = Number(p.CurrentStock) - qty; // update in-memory for the batch write below
      lines.push({ name: p.Name, qty: qty, unitPrice: unitPrice, lineTotal: lineTotal });
    });

    // batch-write updated stock counts
    Object.keys(byId).forEach(function (id) {
      var p = byId[id];
      sheetObjs.sheet.getRange(p.__row, 4).setValue(p.CurrentStock);
      sheetObjs.sheet.getRange(p.__row, 8).setValue(now);
    });

    return { ok: true, saleId: saleId, lines: lines, total: round2(total) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Dashboard & reports ---------------- */

function periodRange(period) {
  var now = new Date();
  var end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1); // exclusive, tomorrow 00:00
  var start;
  if (period === 'week') {
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 6);
  } else if (period === 'month') {
    start = new Date(now.getFullYear(), now.getMonth(), 1);
  } else {
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }
  return { start: start, end: end };
}

function aggregateSales(range) {
  var rows = readSheetObjects(SHEETS.SALES).rows;
  var qtySold = 0, revenue = 0, revenueIfNoDiscount = 0, profit = 0;
  var byProductMap = {}, dailyMap = {};

  rows.forEach(function (r) {
    var ts = new Date(r.Timestamp);
    if (ts >= range.start && ts < range.end) {
      var qty = Number(r.Qty) || 0;
      var unitPrice = Number(r.UnitPriceUsed) || 0;
      var buyPrice = Number(r.BuyPriceAtSale) || 0;
      var noDiscPrice = Number(r.SellPriceNoDiscountAtSale) || 0;
      var lineTotal = Number(r.LineTotal) || (unitPrice * qty);

      qtySold += qty;
      revenue += lineTotal;
      revenueIfNoDiscount += noDiscPrice * qty;
      profit += (unitPrice - buyPrice) * qty;

      var key = r.ProductName;
      if (!byProductMap[key]) byProductMap[key] = { name: key, qty: 0, revenue: 0, profit: 0 };
      byProductMap[key].qty += qty;
      byProductMap[key].revenue += lineTotal;
      byProductMap[key].profit += (unitPrice - buyPrice) * qty;

      var dayKey = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (!dailyMap[dayKey]) dailyMap[dayKey] = { date: dayKey, qty: 0, revenue: 0, profit: 0 };
      dailyMap[dayKey].qty += qty;
      dailyMap[dayKey].revenue += lineTotal;
      dailyMap[dayKey].profit += (unitPrice - buyPrice) * qty;
    }
  });

  var byProduct = Object.keys(byProductMap).map(function (k) { return byProductMap[k]; })
    .sort(function (a, b) { return b.revenue - a.revenue; })
    .map(function (p) { return { name: p.name, qty: p.qty, revenue: round2(p.revenue), profit: round2(p.profit) }; });

  var daily = Object.keys(dailyMap).map(function (k) { return dailyMap[k]; })
    .sort(function (a, b) { return a.date < b.date ? -1 : 1; })
    .map(function (d) { return { date: d.date, qty: d.qty, revenue: round2(d.revenue), profit: round2(d.profit) }; });

  return {
    qtySold: qtySold,
    revenue: round2(revenue),
    revenueIfNoDiscount: round2(revenueIfNoDiscount),
    discountGiven: round2(revenueIfNoDiscount - revenue),
    profit: round2(profit),
    byProduct: byProduct,
    daily: daily
  };
}

function getDashboard() {
  var products = getProductsList();
  var lowStock = products.filter(function (p) { return p.stock < LOW_STOCK_THRESHOLD; });
  return {
    products: products,
    lowStock: lowStock,
    today: aggregateSales(periodRange('day')),
    week: aggregateSales(periodRange('week')),
    month: aggregateSales(periodRange('month'))
  };
}

function getReports(period) {
  var range = periodRange(period);
  var agg = aggregateSales(range);
  agg.period = period;
  agg.rangeStart = Utilities.formatDate(range.start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  agg.rangeEnd = Utilities.formatDate(new Date(range.end.getTime() - 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return agg;
}