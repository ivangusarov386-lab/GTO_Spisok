// Apps Script для таблицы «протокол главный».
// Развёртывание: Развернуть → Управление развёртываниями → ✏️ → Версия: «Новая версия» → Развернуть.
// Доступ: «Выполнять от имени: Я», «У кого есть доступ: Все».

var SHEET_NAME = 'Список';
var COL_NUM = 1;      // A — порядковый номер
var COL_PRESENT = 17; // Q — отметка «пришёл»

function doGet(e) {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonOut({ ok: true, data: [] });
  }
  var values = sheet.getRange(2, 1, lastRow - 1, COL_PRESENT).getValues(); // A:Q
  var data = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var num = row[0]; // A
    var fio = row[1]; // B
    if (num === '' || num === null || fio === '' || fio === null) continue;
    data.push({
      row: i + 2,
      num: num,
      fio: fio,
      uin: row[2],     // C — УИН
      stupen: row[3],  // D — ступень
      pol: row[4],     // E — пол
      present: isChecked(row[COL_PRESENT - 1])
    });
  }
  return jsonOut({ ok: true, data: data });
}

function doPost(e) {
  // Одна запись за раз — несколько телефонов не мешают друг другу
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var params = JSON.parse(e.postData.contents);
    var row = parseInt(params.row, 10);
    var present = !!params.present;
    var sheet = getSheet();

    // Строки могли сдвинуть (сортировка, вставка) — сверяем номер участника
    // и при расхождении ищем его строку по столбцу A.
    if (params.num !== undefined && params.num !== null && params.num !== '') {
      var want = String(params.num).trim();
      var here = row >= 2 ? String(sheet.getRange(row, COL_NUM).getValue()).trim() : '';
      if (here !== want) {
        row = findRowByNum(sheet, want);
        if (!row) return jsonOut({ ok: false, error: 'участник №' + want + ' не найден в листе' });
      }
    }
    if (!row || row < 2) {
      return jsonOut({ ok: false, error: 'некорректный номер строки' });
    }

    var cell = sheet.getRange(row, COL_PRESENT);
    cell.setValue(present);
    SpreadsheetApp.flush();
    var saved = isChecked(cell.getValue());
    if (saved !== present) {
      return jsonOut({ ok: false, error: 'ячейка Q' + row + ' не приняла значение (проверьте флажок/проверку данных)' });
    }
    return jsonOut({ ok: true, row: row, present: saved });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('нет листа «' + SHEET_NAME + '»');
  return sheet;
}

function findRowByNum(sheet, num) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var col = sheet.getRange(2, COL_NUM, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() === num) return i + 2;
  }
  return 0;
}

// Флажок, TRUE/ИСТИНА, «да», 1 — всё считаем отметкой
function isChecked(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'истина' || s === 'да' || s === '1' || s === '✓';
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
