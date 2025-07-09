const SPREADSHEET_ID = '1IIBmI7GzT8aNogdM-3pmROsR1HN1xZASJ0S4AyuKzxA';


// Web App 入口
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

// 查詢客戶訂單 function
function searchOrdersByCustomer(customerName) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('已下訂單');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const indices = {
    date: headers.indexOf('日期'),
    orderNumber: headers.indexOf('訂單編號'),
    customerName: headers.indexOf('客戶名稱'),
    dealerName: headers.indexOf('經銷商名稱'),
    productName: headers.indexOf('設備名稱'),
    quantity: headers.indexOf('數量'),
    status: headers.indexOf('訂單狀態'),
    shippingDate: headers.indexOf('送貨備註')
  };

  const keyword = customerName.toLowerCase();

  const orders = data.slice(1)
    .filter(row =>
      (row[indices.customerName] || '').toLowerCase().includes(keyword) ||
      (row[indices.dealerName] || '').toLowerCase().includes(keyword)
    )
    .map(row => ({
      date: formatDate(row[indices.date]),
      orderNumber: row[indices.orderNumber],
      customerName: row[indices.customerName],
      dealerName: row[indices.dealerName],
      productName: row[indices.productName],
      quantity: row[indices.quantity],
      status: row[indices.status],
      shippingDate: row[indices.shippingDate] ? formatDate(row[indices.shippingDate]) : ''
    }))
    .slice(0, 5);

  if (orders.length === 0) {
    return `❌ 找不到 "${customerName}" 的訂單記錄。`;
  }

  let response = `✅ 找到 ${orders.length} 筆 "${customerName}" 的訂單：\n\n`;
  orders.forEach((order, index) => {
    response += `訂單 ${index + 1}:\n`;
    response += `📅 日期: ${order.date}\n`;
    response += `📋 訂單編號: ${order.orderNumber}\n`;
    response += `👤 客戶: ${order.customerName}\n`;
    response += `🏢 經銷商: ${order.dealerName}\n`;
    response += `📦 產品: ${order.productName} x ${order.quantity}\n`;
    response += `📌 訂單狀態: ${order.status}\n`;
    if (order.shippingDate) {
      response += `🚚 出貨: ${order.shippingDate}\n`;
    }
    response += '\n';
  });

  return response;
}
// 日期格式化
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

