// =============================================
// FMCG Warehouse Dashboard - Server-side Code
// =============================================

var SHEET_ID = '1iKJgXDuw3Msu2grHrZPWJXlzTs5ntAyULuDSNolqUkA';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('FMCG Warehouse Dashboard')
    .setFaviconUrl('https://img.icons8.com/color/48/warehouse.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include external HTML parts
function include(filename) {
  if (!filename || typeof filename !== 'string') {
    return '<!-- include bị gọi sai: tham số không hợp lệ -->';
  }
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return '<!-- Không tìm thấy file: ' + filename + ' -->';
  }
}

// =============================================
// DATA FETCHING (an toàn khi sheet rỗng)
// =============================================

function getSheetData(sheetName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = sheet.getLastColumn();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var data = dataRange.getValues();

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getValues()[0].map(function(h) { return h.toString().trim(); });

  return data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, idx) { obj[header] = row[idx]; });
    return obj;
  });
}

// =============================================
// AGGREGATION FUNCTIONS
// =============================================

function getInventoryOverview() {
  var data = getSheetData('Tồn Kho');
  var totalValue = data.reduce(function(sum, r) { return sum + (parseFloat(r['Giá Trị Tồn (VNĐ)']) || 0); }, 0);
  var totalSKU = data.length;
  var lowStock = data.filter(function(r) { return (parseFloat(r['Tồn Kho']) || 0) < (parseFloat(r['Tồn Tối Thiểu']) || 0); });
  var lowStockCount = lowStock.length;

  var categoryValues = {};
  data.forEach(function(r) {
    var cat = r['Danh Mục'] || 'Khác';
    categoryValues[cat] = (categoryValues[cat] || 0) + (parseFloat(r['Giá Trị Tồn (VNĐ)']) || 0);
  });

  var sorted = data.slice().sort(function(a,b) { return (parseFloat(b['Giá Trị Tồn (VNĐ)'])||0) - (parseFloat(a['Giá Trị Tồn (VNĐ)'])||0); });
  var top10Highest = sorted.slice(0,10).map(function(r) {
    return { sku: r['Mã SKU'], name: r['Tên Hàng'], value: r['Giá Trị Tồn (VNĐ)'], stock: r['Tồn Kho'], minStock: r['Tồn Tối Thiểu'] };
  });
  var top10Lowest = sorted.slice(-10).reverse().map(function(r) {
    return { sku: r['Mã SKU'], name: r['Tên Hàng'], value: r['Giá Trị Tồn (VNĐ)'], stock: r['Tồn Kho'], minStock: r['Tồn Tối Thiểu'] };
  });

  // ABC analysis
  var total = totalValue;
  var cumulative = 0;
  var aItems = 0, bItems = 0, cItems = 0;
  sorted.forEach(function(r) {
    cumulative += (parseFloat(r['Giá Trị Tồn (VNĐ)']) || 0);
    var pct = cumulative / total;
    if (pct <= 0.7) aItems++;
    else if (pct <= 0.9) bItems++;
    else cItems++;
  });

  return {
    totalValue: totalValue,
    totalSKU: totalSKU,
    lowStockCount: lowStockCount,
    categoryValues: categoryValues,
    top10Highest: top10Highest,
    top10Lowest: top10Lowest,
    abc: { a: aItems, b: bItems, c: cItems }
  };
}

function getInboundOutboundSummary() {
  var inbound = getSheetData('Nhập Kho');
  var outbound = getSheetData('Xuất Kho');

  var inboundByMonth = {};
  inbound.forEach(function(r) {
    var d = new Date(r['Ngày']);
    var key = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    inboundByMonth[key] = (inboundByMonth[key] || 0) + (parseFloat(r['Số Lượng']) || 0);
  });
  var outboundByMonth = {};
  outbound.forEach(function(r) {
    var d = new Date(r['Ngày']);
    var key = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    outboundByMonth[key] = (outboundByMonth[key] || 0) + (parseFloat(r['Số Lượng']) || 0);
  });

  var supplierTotals = {};
  inbound.forEach(function(r) {
    var sup = r['Nhà Cung Cấp'] || 'Unknown';
    supplierTotals[sup] = (supplierTotals[sup] || 0) + (parseFloat(r['Thành Tiền (VNĐ)']) || 0);
  });
  var topSuppliers = Object.entries(supplierTotals).sort(function(a,b){ return b[1]-a[1]; }).slice(0,10).map(function(e){ return {name: e[0], value: e[1]}; });

  var channelSales = {};
  outbound.forEach(function(r) {
    var ch = r['Kênh Bán'] || 'Unknown';
    channelSales[ch] = (channelSales[ch] || 0) + (parseFloat(r['Thành Tiền (VNĐ)']) || 0);
  });

  var customerSales = {};
  outbound.forEach(function(r) {
    var cust = r['Khách Hàng'] || 'Unknown';
    customerSales[cust] = (customerSales[cust] || 0) + (parseFloat(r['Thành Tiền (VNĐ)']) || 0);
  });
  var topCustomers = Object.entries(customerSales).sort(function(a,b){ return b[1]-a[1]; }).slice(0,10).map(function(e){ return {name: e[0], value: e[1]}; });

  // Pareto (20% SKU)
  var skuSales = {};
  outbound.forEach(function(r) {
    var sku = r['Mã SKU'];
    skuSales[sku] = (skuSales[sku] || 0) + (parseFloat(r['Thành Tiền (VNĐ)']) || 0);
  });
  var sortedSku = Object.entries(skuSales).sort(function(a,b){ return b[1]-a[1]; });
  var totalSales = sortedSku.reduce(function(s, e){ return s+e[1]; }, 0);
  var pareto20 = Math.ceil(sortedSku.length * 0.2);
  var paretoValue = 0;
  for (var i=0; i<pareto20; i++) {
    paretoValue += sortedSku[i][1];
  }

  return {
    inboundByMonth: inboundByMonth,
    outboundByMonth: outboundByMonth,
    topSuppliers: topSuppliers,
    channelSales: channelSales,
    topCustomers: topCustomers,
    pareto: { skuCount20: pareto20, valuePct: totalSales ? (paretoValue/totalSales*100).toFixed(1) : 0 }
  };
}

function getStockAccuracy() {
  var kiemKe = getSheetData('Kiểm Kê');
  var dieuChinh = getSheetData('Điều Chỉnh');

  var totalAbsoluteDiff = 0, totalCheckedValue = 0;
  kiemKe.forEach(function(r) {
    var diff = Math.abs(parseFloat(r['Chênh Lệch']) || 0);
    totalAbsoluteDiff += diff;
    totalCheckedValue += parseFloat(r['Thực Tế Đếm']) || 0;
  });
  var discrepancyRate = totalCheckedValue ? (totalAbsoluteDiff / totalCheckedValue * 100).toFixed(2) : 0;

  var monthlyChecks = {};
  kiemKe.forEach(function(r) {
    var d = new Date(r['Ngày Kiểm']);
    var key = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    if (!monthlyChecks[key]) monthlyChecks[key] = { total:0, withDiff:0 };
    monthlyChecks[key].total++;
    if (parseFloat(r['Chênh Lệch']) !== 0) monthlyChecks[key].withDiff++;
  });

  var reasonCount = {};
  dieuChinh.forEach(function(r) {
    var reason = r['Lý Do'] || r['Loại Điều Chỉnh'] || 'Khác';
    reasonCount[reason] = (reasonCount[reason] || 0) + Math.abs(parseFloat(r['Số Lượng']) || 0);
  });

  return {
    discrepancyRate: discrepancyRate,
    monthlyChecks: monthlyChecks,
    reasonCount: reasonCount
  };
}

function getFinancialSummary() {
  var inventory = getSheetData('Tồn Kho');
  var outbound = getSheetData('Xuất Kho');

  var totalRevenue = outbound.reduce(function(s, r){ return s + (parseFloat(r['Thành Tiền (VNĐ)']) || 0); }, 0);
  var invMap = {};
  inventory.forEach(function(r){ invMap[r['Mã SKU']] = parseFloat(r['Giá Vốn (VNĐ)']) || 0; });
  var totalCOGS = outbound.reduce(function(s, r){
    var qty = parseFloat(r['Số Lượng']) || 0;
    var cost = invMap[r['Mã SKU']] || 0;
    return s + qty * cost;
  }, 0);
  var grossProfit = totalRevenue - totalCOGS;
  var grossMargin = totalRevenue ? (grossProfit / totalRevenue * 100).toFixed(1) : 0;

  var totalInventoryValue = inventory.reduce(function(s, r){ return s + (parseFloat(r['Giá Trị Tồn (VNĐ)']) || 0); }, 0);
  var avgInventory = totalInventoryValue;
  var inventoryTurnover = avgInventory ? (totalCOGS / avgInventory) : 0;
  var daysInventory = inventoryTurnover ? (365 / inventoryTurnover).toFixed(0) : 'N/A';

  return {
    totalRevenue: totalRevenue,
    totalCOGS: totalCOGS,
    grossProfit: grossProfit,
    grossMargin: grossMargin,
    totalInventoryValue: totalInventoryValue,
    inventoryTurnover: inventoryTurnover.toFixed(2),
    daysInventory: daysInventory
  };
}

function getGenZInsights() {
  var outbound = getSheetData('Xuất Kho');

  var onlineByMonth = {};
  var totalByMonth = {};
  outbound.forEach(function(r) {
    var d = new Date(r['Ngày']);
    var key = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    totalByMonth[key] = (totalByMonth[key] || 0) + (parseFloat(r['Thành Tiền (VNĐ)']) || 0);
    if (r['Kênh Bán'] === 'Online') {
      onlineByMonth[key] = (onlineByMonth[key] || 0) + (parseFloat(r['Thành Tiền (VNĐ)']) || 0);
    }
  });

  var onlineShare = Object.keys(totalByMonth).map(function(month) {
    return {
      month: month,
      online: onlineByMonth[month] || 0,
      total: totalByMonth[month],
      share: totalByMonth[month] ? ((onlineByMonth[month]||0)/totalByMonth[month]*100).toFixed(1) : 0
    };
  });

  // Promotion rate by channel
  var promoCounts = {};
  var orderCounts = {};
  var orders = {};
  outbound.forEach(function(r) {
    var ch = r['Kênh Bán'] || 'Other';
    var orderId = r['Mã Phiếu'];
    if (!orders[orderId]) {
      orders[orderId] = { channels: new Set(), hasPromo: false };
      orderCounts[ch] = (orderCounts[ch] || 0) + 1;
    }
    orders[orderId].channels.add(ch);
    if (r['Ghi Chú'] && r['Ghi Chú'].toString().toLowerCase().indexOf('khuyến mãi') > -1) {
      orders[orderId].hasPromo = true;
    }
  });

  Object.values(orders).forEach(function(order) {
    order.channels.forEach(function(ch) {
      if (order.hasPromo) promoCounts[ch] = (promoCounts[ch] || 0) + 1;
    });
  });

  var promoRate = {};
  Object.keys(orderCounts).forEach(function(ch) {
    promoRate[ch] = {
      total: orderCounts[ch],
      promos: promoCounts[ch] || 0,
      rate: orderCounts[ch] ? ((promoCounts[ch]||0)/orderCounts[ch]*100).toFixed(1) : 0
    };
  });

  var onlineProducts = {};
  outbound.forEach(function(r) {
    if (r['Kênh Bán'] === 'Online') {
      onlineProducts[r['Tên Hàng']] = (onlineProducts[r['Tên Hàng']] || 0) + (parseFloat(r['Số Lượng']) || 0);
    }
  });
  var topOnline = Object.entries(onlineProducts).sort(function(a,b){ return b[1]-a[1]; }).slice(0,10).map(function(e){ return {name: e[0], qty: e[1]}; });

  return {
    onlineShare: onlineShare,
    promoRate: promoRate,
    topOnline: topOnline
  };
}

function getAllDashboardData() {
  return {
    inventory: getInventoryOverview(),
    logistics: getInboundOutboundSummary(),
    accuracy: getStockAccuracy(),
    finance: getFinancialSummary(),
    genz: getGenZInsights()
  };
}

// Optional: Email alert & trigger setup
function sendLowStockAlert() {
  var data = getSheetData('Tồn Kho');
  var lowStock = data.filter(function(r){ return (parseFloat(r['Tồn Kho'])||0) < (parseFloat(r['Tồn Tối Thiểu'])||0); });
  if (lowStock.length === 0) return;
  // ... send email
}
