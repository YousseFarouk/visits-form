var DAY_NAMES = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
var VISITS_SHEET = 'زيارات المندوبين';
var CLIENTS_SHEET = 'بيانات العملاء';
// ترتيب أعمدة شيت العملاء: A=اليوم, B=اسم العميل, C=التصنيف, D=العنوان, E=رقم التليفون
// ترتيب أعمدة شيت الزيارات: A=التاريخ, B=اسم العميل, C=التصنيف, D=حالة الزيارة, E=خط العرض, F=خط الطول, G=رابط الموقع, H=سبب عدم الزيارة, I=رقم التليفون

// ============================================================
// نقطة الدخول الوحيدة - بترجع JSON بس، مفيش HTML خالص
// كل الأكشنز بتتبعت GET عشان نتفادى مشاكل CORS preflight
// اللي بتحصل مع الطلبات من دومين تاني (زي GitHub Pages)
// ============================================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  var result;

  try {
    switch (action) {

      case 'getTodayClients':
        result = { ok: true, data: getTodayClients() };
        break;

      case 'getClientDetails':
        result = { ok: true, data: getClientDetails(e.parameter.name) };
        break;

      case 'updateClientDetails':
        updateClientDetails(e.parameter.name, e.parameter.phone, e.parameter.address, e.parameter.category);
        result = { ok: true };
        break;

      case 'addNewClient':
        var addMsg = addNewClient(e.parameter.name, e.parameter.phone, e.parameter.address, e.parameter.category);
        result = { ok: addMsg.indexOf('❌') !== 0, message: addMsg };
        break;

      case 'saveVisit':
        var saveMsg = saveVisit(
          e.parameter.name,
          e.parameter.status,
          parseFloat(e.parameter.lat),
          parseFloat(e.parameter.lng),
          e.parameter.reason || '',
          e.parameter.phone,
          e.parameter.category
        );
        result = { ok: true, message: saveMsg };
        break;

      default:
        result = { ok: false, error: 'أكشن غير معروف: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// الدوال الأصلية - نفس المنطق بالظبط، من غير أي تغيير
// ============================================================

function getTodayClients() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET);
  var todayName = DAY_NAMES[new Date().getDay()];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var names = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] && data[i][0] === todayName) { names.push(data[i][1]); }
  }
  return names;
}

function getClientDetails(clientName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === clientName) {
      return { category: data[i][2], address: data[i][3], phone: data[i][4] };
    }
  }
  return null;
}

function updateClientDetails(clientName, phone, address, category) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === clientName) {
      sheet.getRange(i + 2, 3, 1, 3).setValues([[category, address, phone]]);
      return;
    }
  }
}

// بتضيف عميل جديد، بعد ما تتأكد إن رقم تليفونه مش مكرر عند حد تاني
function addNewClient(clientName, phone, address, category) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET);
  var todayName = DAY_NAMES[new Date().getDay()];
  var cleanPhone = phone.toString().replace(/\s+/g, '');

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existingPhones = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    for (var i = 0; i < existingPhones.length; i++) {
      var existing = existingPhones[i][0].toString().replace(/\s+/g, '');
      if (existing !== '' && existing === cleanPhone) {
        return '❌ الرقم ده موجود بالفعل عند عميل تاني، اتأكد منه';
      }
    }
  }

  sheet.appendRow([todayName, clientName, category, address, cleanPhone]);
  return 'تم إضافة العميل ✅';
}

// بيرجع رابط خريطة جاهز على مكان المندوب وقت الحفظ
function reverseGeocode(lat, lng) {
  return 'https://www.google.com/maps?q=' + lat + ',' + lng;
}

// بيدور على آخر صف فيه بيانات فعلية في عمود B
function findLastDataRow(sheet) {
  var colB = sheet.getRange('B2:B').getValues();
  var lastRow = 1;
  for (var i = 0; i < colB.length; i++) {
    if (colB[i][0] !== '') { lastRow = i + 2; }
  }
  return lastRow;
}

function saveVisit(clientName, visitStatus, lat, lng, reason, phone, category) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VISITS_SHEET);
  var now = new Date();
  var mapLink = reverseGeocode(lat, lng);
  var todayName = DAY_NAMES[now.getDay()];

  var lastRow = findLastDataRow(sheet);
  var hasData = !(lastRow === 1 && sheet.getRange('B2').getValue() === '');
  var targetRow = lastRow + 1;

  if (hasData) {
    var lastDateValue = sheet.getRange(lastRow, 1).getValue();
    if (lastDateValue instanceof Date) {
      var lastDayName = DAY_NAMES[lastDateValue.getDay()];
      if (lastDayName !== todayName) {
        targetRow = lastRow + 2; // فاصل بين يومين مختلفين
      }
    }
  } else {
    targetRow = 2;
  }

  sheet.getRange(targetRow, 1, 1, 9).setValues([[now, clientName, category, visitStatus, lat, lng, mapLink, reason, phone]]);
  return 'تم الحفظ بنجاح ✅';
}
