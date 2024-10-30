// نکات
// قیمت ها به تومان وارد شوند


// تنظیمات اجرای مرحله‌ای
const BATCH_SIZE = 100; // تعداد سطرهای پردازش در هر مرحله
const REQUEST_DELAY = 1000; // تاخیر 1 ثانیه بین هر درخواست
const SCRIPT_PROPERTY_KEYS = {
  CURRENT_ROW: 'CURRENT_ROW',
  TOTAL_ROWS: 'TOTAL_ROWS',
  SYNC_IN_PROGRESS: 'SYNC_IN_PROGRESS',
  SYNC_TYPE: 'SYNC_TYPE'
};

var menu = null;

function onOpen() {
  updateMenu();
}

function updateMenu() {
  var ui = SpreadsheetApp.getUi();
  menu = ui.createMenu('همگام‌سازی محصولات باسلام');
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const syncInProgress = scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.SYNC_IN_PROGRESS) === 'true';
  
  menu.addItem('بروزرسانی قیمت و موجودی', 'startUpdateProductsBoth')
    .addItem('بروزرسانی موجودی', 'startUpdateProductsStock')
    .addItem('بروزرسانی قیمت', 'startUpdateProductsPrice')
    .addItem('تنظیم توکن', 'showApiKeyDialog');
  
  if (syncInProgress) {
    menu.addItem('توقف همگام‌سازی', 'stopSync');
  }
  
  menu.addToUi();
}

function showApiKeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'تنظیم توکن باسلام پلاس',
    'توکن باسلام پلاس خود را وارد نمایید:',
    ui.ButtonSet.OK_CANCEL
  );

  var button = result.getSelectedButton();
  var apiKey = result.getResponseText();
  
  if (button == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('BASALAM_PLUS_API_TOKEN', apiKey);
    ui.alert('توکن با موفقیت تنظیم شد.');
  }
}

function startUpdateProductsBoth() {
  startBatchProcessing('stockPrice');
}

function startUpdateProductsStock() {
  startBatchProcessing('stock');
}

function startUpdateProductsPrice() {
  startBatchProcessing('price');
}

function startBatchProcessing(type) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // بررسی وجود ستون‌های مورد نیاز
  if (!validateColumns(data[0], type)) {
    return;
  }
  
  // تنظیم مقادیر اولیه
  scriptProperties.setProperties({
    [SCRIPT_PROPERTY_KEYS.CURRENT_ROW]: '1',
    [SCRIPT_PROPERTY_KEYS.TOTAL_ROWS]: data.length.toString(),
    [SCRIPT_PROPERTY_KEYS.SYNC_IN_PROGRESS]: 'true',
    [SCRIPT_PROPERTY_KEYS.SYNC_TYPE]: type
  });
  
  // شروع پردازش مرحله‌ای
  processBatch();
}

function processBatch() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentRow = parseInt(scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.CURRENT_ROW));
  const totalRows = parseInt(scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.TOTAL_ROWS));
  const type = scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.SYNC_TYPE);
  
  if (currentRow >= totalRows || scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.SYNC_IN_PROGRESS) !== 'true') {
    finishProcessing();
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const endRow = Math.min(currentRow + BATCH_SIZE, totalRows);
  
  for (let i = currentRow; i < endRow; i++) {
    processRow(sheet, data, headers, i, type);
    
    // اضافه کردن تاخیر بین درخواست‌ها
    if (i < endRow - 1) { // برای آخرین آیتم تاخیر نمی‌خواهیم
      Utilities.sleep(REQUEST_DELAY);
    }
  }
  
  // تنظیم سطر بعدی و برنامه‌ریزی اجرای بعدی
  scriptProperties.setProperty(SCRIPT_PROPERTY_KEYS.CURRENT_ROW, endRow.toString());
  
  // زمان‌بندی اجرای بعدی با تاخیر
  ScriptApp.newTrigger('processBatch')
    .timeBased()
    .after(REQUEST_DELAY)
    .create();
}

function processRow(sheet, data, headers, rowIndex, type) {
  const idIndex = headers.indexOf('شناسه محصول');
  const priceIndex = headers.indexOf('قیمت');
  const skuIndex = headers.indexOf('شناسه داخلی');
  const stockIndex = headers.indexOf('موجودی');
  const colorIndex = headers.indexOf('رنگ');
  const sizeIndex = headers.indexOf('سایز');
  const resultIndex = headers.indexOf('نتیجه بروزرسانی') === -1 ? 
    headers.length : headers.indexOf('نتیجه بروزرسانی');
  
  let result;
  try {
    const row = data[rowIndex];
    const requestData = prepareRequestData(row, type, {
      idIndex, priceIndex, stockIndex, skuIndex, colorIndex, sizeIndex
    });
    
    if (row[colorIndex] && row[sizeIndex]) {
      result = {
        result: false,
        message: "نمیتوانید همزمان هم رنگ و هم سایز در یک ردیف وارد کنید"
      };
    } else {
      setCellColor(sheet, rowIndex, idIndex, 'processing');
      result = basalamPlusRequester('/v2.0/user/product/edit', requestData);
      if (result.result && typeof result.message === 'undefined') {
        result.message = "انجام شد";
      } else if (result.result) {
        result.message = "انجام شد";
      }
    }
  } catch (error) {
    result = {
      result: false,
      message: `خطا: ${error.message}`
    };
    console.error(`Error processing row ${rowIndex + 1}:`, error);
  }
  
  // بروزرسانی نتیجه در شیت
  updateResultInSheet(sheet, rowIndex, idIndex, resultIndex, result);
}

function prepareRequestData(row, type, indices) {
  const { idIndex, priceIndex, stockIndex, skuIndex, colorIndex, sizeIndex } = indices;
  
  let requestData = {
    id: row[idIndex]
  };
  
  switch(type) {
    case "stockPrice":
      requestData.price = row[priceIndex] * 10;
      requestData.stock = row[stockIndex];
      break;
    case "stock":
      requestData.stock = row[stockIndex];
      break;
    case "price":
      requestData.price = row[priceIndex] * 10;
      break;
  }
  
  if (row[skuIndex]) {
    requestData.sku = row[skuIndex];
  }
  
  // اضافه کردن variant در صورت وجود رنگ یا سایز
  if (row[colorIndex] || row[sizeIndex]) {
    const variant = {};
    if (row[colorIndex]) {
      variant.property = "color";
      variant.value = row[colorIndex];
    } else if (row[sizeIndex]) {
      variant.property = "size";
      variant.value = row[sizeIndex];
    }
    
    // کپی کردن سایر مقادیر به variant
    Object.keys(requestData).forEach(key => {
      if (key !== 'variant' && requestData[key]) {
        variant[key] = requestData[key];
      }
    });
    
    requestData.variant = variant;
  }
  
  return requestData;
}

function updateResultInSheet(sheet, rowIndex, idIndex, resultIndex, result) {
  if (result.result) {
    setCellColor(sheet, rowIndex, idIndex, 'success');
    setCellColor(sheet, rowIndex, resultIndex, 'success');
  } else {
    setCellColor(sheet, rowIndex, idIndex, 'failure');
    setCellColor(sheet, rowIndex, resultIndex, 'failure');
  }
  
  sheet.getRange(rowIndex + 1, resultIndex + 1).setValue(result.message);
}

function finishProcessing() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    [SCRIPT_PROPERTY_KEYS.SYNC_IN_PROGRESS]: 'false',
    [SCRIPT_PROPERTY_KEYS.CURRENT_ROW]: '',
    [SCRIPT_PROPERTY_KEYS.TOTAL_ROWS]: '',
    [SCRIPT_PROPERTY_KEYS.SYNC_TYPE]: ''
  });
  
  updateMenu();
  SpreadsheetApp.getUi().alert('بروزرسانی محصولات به پایان رسید.');
  
  // حذف همه تریگرهای زمان‌بندی شده
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function stopSync() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_PROPERTY_KEYS.SYNC_IN_PROGRESS, 'false');
  SpreadsheetApp.getUi().alert('همگام‌سازی متوقف شد.');
  updateMenu();
}

function validateColumns(headers, type) {
  const ui = SpreadsheetApp.getUi();
  const idIndex = headers.indexOf('شناسه محصول');
  const priceIndex = headers.indexOf('قیمت');
  const stockIndex = headers.indexOf('موجودی');
  
  if (idIndex === -1) {
    ui.alert('ستون "شناسه محصول" یافت نشد.');
    return false;
  }
  
  switch(type) {
    case "stockPrice":
      if (priceIndex === -1 || stockIndex === -1) {
        ui.alert('ستون‌های مورد نیاز یافت نشد. لطفاً مطمئن شوید که ستون‌های "قیمت" و "موجودی" وجود دارند.');
        return false;
      }
      break;
    case "stock":
      if (stockIndex === -1) {
        ui.alert('لطفاً مطمئن شوید که ستون‌ "موجودی" وجود دارد.');
        return false;
      }
      break;
    case "price":
      if (priceIndex === -1) {
        ui.alert('لطفاً مطمئن شوید که ستون‌ "قیمت" وجود دارد.');
        return false;
      }
      break;
  }
  
  return true;
}

function setCellColor(sheet, row, col, status) {
  var cell = sheet.getRange(row + 1, col + 1);
  switch(status) {
    case 'processing':
      cell.setBackground('#FFFDE7'); // Light yellow
      break;
    case 'success':
      cell.setBackground('#E8F5E9'); // Light green
      break;
    case 'failure':
      cell.setBackground('#FFEBEE'); // Light red
      break;
    default:
      cell.setBackground('#FFFFFF'); // White (reset)
  }
}

function basalamPlusRequester(uri, data, method="POST") {
  var apiBaseUrl = 'https://plus.basalam.com/api';
  var apiUrl = apiBaseUrl + uri;
  var apiKey = PropertiesService.getScriptProperties().getProperty('BASALAM_PLUS_API_TOKEN');

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("توکن تنظیم نشده");
    throw new Error("توکن تنظیم نشده");
  }
  
  var headers = {
    'Authorization': 'Bearer ' + apiKey.trim(),
    'Content-Type': 'application/json'
  };
  
  var options = {
    'method': method,
    'headers': headers,
    'payload': JSON.stringify(data),
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  
  var result;
  try {
    responseText = decodeUnicodeEscapes(responseText);
    result = JSON.parse(responseText);
    result = recursivelyDecodeUnicode(result);
  } catch (e) {
    result = { result: false, message: "خطا در تجزیه پاسخ", data: responseText };
  }
  
  result.httpStatus = responseCode;
  
  if (responseCode < 200 || responseCode >= 300) {
    result.result = false;
    if (!result.message) {
      result.message = "خطای HTTP: " + responseCode;
    }
  }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// تابع کمکی برای تبدیل کدهای یونیکد به کاراکترهای فارسی
function decodeUnicodeEscapes(str) {
  return str.replace(/\\u([\d\w]{4})/gi, function (match, grp) {
    return String.fromCharCode(parseInt(grp, 16));
  });
}

// تابع کمکی برای تبدیل بازگشتی تمام مقادیر رشته‌ای درون یک آبجکت
function recursivelyDecodeUnicode(obj) {
  if (typeof obj === 'string') {
    return decodeUnicodeEscapes(obj);
  } else if (Array.isArray(obj)) {
    return obj.map(recursivelyDecodeUnicode);
  } else if (typeof obj === 'object' && obj !== null) {
    Object.keys(obj).forEach(key => {
      obj[key] = recursivelyDecodeUnicode(obj[key]);
    });
  }
  return obj;
}
