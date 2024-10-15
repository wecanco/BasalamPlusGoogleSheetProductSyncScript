// نکات
// قیمت ها به تومان وارد شود


var menu = null;
let syncInProgress = false;

function onOpen() {
  updateMenu();
}

function updateMenu() {
  var ui = SpreadsheetApp.getUi();
  menu = ui.createMenu('همگام‌سازی محصولات باسلام');
  menu.addItem('بروزرسانی قیمت و موجودی', 'updateProductsBoth')
    .addItem('بروزرسانی موجودی', 'updateProductsStock')
    .addItem('بروزرسانی قیمت', 'updateProductsPrice')
    // .addItem('افزودن محصولات', 'addProducts')
    // .addItem('بروزرسانی کامل محصولات', 'editProducts')
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
  } else {
    //ui.alert('تنظیم توکن لغو شد.');
  }
}

function updateProductsBoth() {
  updateProducts('stockPrice');
}
function updateProductsStock() {
  updateProducts('stock');
}
function updateProductsPrice() {
  updateProducts('price');
}


function updateProducts(type=null) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idIndex = headers.indexOf('شناسه محصول');
  var priceIndex = headers.indexOf('قیمت');
  var skuIndex = headers.indexOf('شناسه داخلی');
  var stockIndex = headers.indexOf('موجودی');
  var colorIndex = headers.indexOf('رنگ');
  var sizeIndex = headers.indexOf('سایز');
  var resultIndex = headers.indexOf('نتیجه بروزرسانی');
  
  if (idIndex === -1) {
    SpreadsheetApp.getUi().alert('ستون "شناسه محصول" یافت نشد.');
    return;
  }

  syncInProgress = true;
  updateMenu();
  var requestData = null;

  switch(type) {
    case "stockPrice":
      if (priceIndex === -1 || stockIndex === -1) {
        SpreadsheetApp.getUi().alert('ستون‌های مورد نیاز یافت نشد. لطفاً مطمئن شوید که ستون‌های "قیمت" و "موجودی" وجود دارند.');
        syncInProgress = false;
        updateMenu();
        return;
      }
    break;

    case "stock":
      if (stockIndex === -1) {
        SpreadsheetApp.getUi().alert('لطفاً مطمئن شوید که ستون‌ "موجودی" وجود دارد.');
        syncInProgress = false;
        updateMenu();
        return;
      }
      
    break;

    case "price":
      if (priceIndex === -1) {
        SpreadsheetApp.getUi().alert('لطفاً مطمئن شوید که ستون‌ "قیمت" وجود دارد.');
        syncInProgress = false;
        updateMenu();
        return;
      }
      
    break;

  }
  
  if (resultIndex === -1) {
    resultIndex = headers.length;
    sheet.getRange(1, resultIndex + 1).setValue('نتیجه بروزرسانی');
  }

  for (var i = 1; i < data.length; i++) {
    var productId = data[i][idIndex];
    var price = data[i][priceIndex] * 10;
    var stock = data[i][stockIndex];
    var color = data[i][colorIndex];
    var size = data[i][sizeIndex];
    var sku = data[i][skuIndex];

    switch(type) {
      case "stockPrice":
        requestData = {
          "id": productId,
          "price": price,
          "stock": stock,
        };
      break;
      case "stock":
        requestData = {
          "id": productId,
          "stock": stock
        };
      break;
      case "price":
        requestData = {
          "id": productId,
          "price": price
        };
      break;
    }

    if (sku) {
      requestData['sku'] = sku;
    }

    var _result = null;

    if (color && size) {
      _result = {"result": false, "message": "نمیتوانید همزمان هم رنگ و هم سایز در یک ردیف وارد کنید", "data": null};
    } else {
      if (color || size) {
        requestData['variant'] = null;

        if (color) {
          variant = {
              "property": "color",
              "value": color,
          };
          
        }

        if (size) {
          variant = {
              "property": "size",
              "value": size,
          };
          
        }

        for (const [key, value] of Object.entries(requestData)) {
          if (key != 'variant' && value) {
            variant[key] = value;
          }
        }

        requestData['variant'] = variant;

      }

      setCellColor(sheet, i, idIndex, 'processing');

      _result = basalamPlusRequester('/v2.0/user/product/edit', requestData)
      if (_result.result && typeof _result.message !== 'undefined') {
        _result.message = "انجام شد";
      }
    }

    if (_result.result) {
      setCellColor(sheet, i, idIndex, 'success');
      setCellColor(sheet, i, resultIndex, 'success');
    } else {
      setCellColor(sheet, i, idIndex, 'failure');
      setCellColor(sheet, i, resultIndex, 'failure');
    }
    
    sheet.getRange(i + 1, resultIndex + 1).setValue(_result.message);
    
  }

  syncInProgress = false;
  updateMenu();
  SpreadsheetApp.getUi().alert('بروزرسانی محصولات به پایان رسید.');
}

function basalamPlusRequester(uri, data, method="POST") {
  var apiBaseUrl = 'https://plus.basalam.com/api';
  var apiUrl = apiBaseUrl+uri;
  var apiKey = PropertiesService.getScriptProperties().getProperty('BASALAM_PLUS_API_TOKEN');

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("توکن تنظیم نشده");
    syncInProgress = false;
    updateMenu();
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
    // تبدیل رشته JSON با کدهای یونیکد به متن فارسی قابل خواندن
    responseText = decodeUnicodeEscapes(responseText);
    result = JSON.parse(responseText);
    
    // تبدیل بازگشتی تمام مقادیر رشته‌ای درون آبجکت نتیجه
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
