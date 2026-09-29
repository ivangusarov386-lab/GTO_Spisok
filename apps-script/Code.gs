// ===== НАЧАЛО ФАЙЛА. Вставляйте весь файл целиком (около 200 строк), до метки «КОНЕЦ ФАЙЛА». =====
// Apps Script для таблицы «протокол главный».
// Развёртывание: Развернуть → Управление развёртываниями → ✏️ → Версия: «Новая версия» → Развернуть.
// Доступ: «Выполнять от имени: Я», «У кого есть доступ: Все».

var SHEET_NAME = 'Список';
var COL_NUM = 1;      // A — порядковый номер
var COL_FIO = 2;      // B — ФИО
var COL_UIN = 3;      // C — УИН
var COL_STUPEN = 4;   // D — ступень
var COL_POL = 5;      // E — пол
var COL_PRESENT = 17; // Q — отметка «пришёл»
var COL_RES_FIRST = 6;  // F — первый норматив (отжимание)
var COL_RES_LAST = 16;  // P — последний норматив (60м)
var COL_CLASS = 18;     // R — класс (цифра)
var COL_LETTER = 19;    // S — буква класса

function doGet(e) {
  if (e && e.parameter && e.parameter.view === 'results') return jsonOut(getResults());
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
    if (params.action === 'add') return jsonOut(addParticipant(params));
    if (params.action && params.action !== 'mark') {
      return jsonOut({ ok: false, error: 'неизвестное действие: ' + params.action });
    }
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

// Новый участник: строка сразу после последнего заполненного ФИО,
// номер = максимальный в столбце A + 1, оформление — как у строки выше.
function addParticipant(p) {
  var fio = String(p.fio || '').replace(/\s+/g, ' ').trim();
  var pol = String(p.pol || '').trim();
  var uin = String(p.uin || '').trim();
  var stupen = String(p.stupen || '').trim();
  var uinDigits = uin.replace(/\D/g, '');
  if (fio.split(' ').length < 2) return { ok: false, error: 'нужны фамилия и имя' };
  if (!pol) return { ok: false, error: 'не указан пол' };
  if (uinDigits.length !== 11) return { ok: false, error: 'УИН должен содержать 11 цифр' };
  if (!stupen) return { ok: false, error: 'не указана ступень' };

  var sheet = getSheet();
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var maxNum = 0, lastFioRow = 1;
  if (lastRow >= 2) {
    var vals = sheet.getRange(2, 1, lastRow - 1, COL_POL).getValues(); // A:E
    for (var i = 0; i < vals.length; i++) {
      var n = parseInt(vals[i][COL_NUM - 1], 10);
      if (n > maxNum) maxNum = n;
      if (String(vals[i][COL_FIO - 1]).trim() !== '') lastFioRow = i + 2;
      // Защита от двойного добавления (повтор после обрыва связи)
      if (String(vals[i][COL_UIN - 1]).replace(/\D/g, '') === uinDigits) {
        return { ok: false, error: 'УИН уже есть в таблице: №' + vals[i][COL_NUM - 1] + ' ' + vals[i][COL_FIO - 1] };
      }
    }
  }
  var num = maxNum + 1;
  var newRow = lastFioRow + 1;
  if (newRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);

  // Оформление (шрифты, рамки, форматы) и флажок в Q — как у строки выше
  if (lastFioRow >= 2) {
    var lastCol = sheet.getLastColumn();
    sheet.getRange(lastFioRow, 1, 1, lastCol).copyTo(sheet.getRange(newRow, 1, 1, lastCol), { formatOnly: true });
    var dv = sheet.getRange(lastFioRow, COL_PRESENT).getDataValidation();
    if (dv) sheet.getRange(newRow, COL_PRESENT).setDataValidation(dv);
  }
  sheet.getRange(newRow, COL_UIN).setNumberFormat('@'); // УИН — текстом, чтобы не стал числом
  sheet.getRange(newRow, COL_NUM, 1, 5).setValues([[num, fio, uin, stupen, pol]]); // A:E
  sheet.getRange(newRow, COL_PRESENT).setValue(false);
  SpreadsheetApp.flush();
  return { ok: true, row: newRow, num: num };
}

// Витрина результатов: заголовки нормативов из строки 1 и значения ровно
// как они видны в таблице (getDisplayValues — «12,10» не превратится в 12.1)
function getResults() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = Math.min(COL_LETTER, sheet.getMaxColumns()); // A:S
  var head = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var disciplines = [];
  for (var c = COL_RES_FIRST; c <= COL_RES_LAST; c++) {
    disciplines.push({ col: columnLetter(c), name: String(head[c - 1] || columnLetter(c)).trim() });
  }
  var data = [];
  if (lastRow >= 2) {
    var shown = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    var raw = sheet.getRange(2, COL_PRESENT, lastRow - 1, 1).getValues();
    for (var i = 0; i < shown.length; i++) {
      var r = shown[i];
      if (String(r[0]).trim() === '' || String(r[1]).trim() === '') continue;
      data.push({
        row: i + 2,
        num: r[COL_NUM - 1],
        fio: r[COL_FIO - 1],
        uin: r[COL_UIN - 1],
        stupen: r[COL_STUPEN - 1],
        pol: r[COL_POL - 1],
        present: isChecked(raw[i][0]),
        klass: String(r[COL_CLASS - 1] || '').trim(),
        letter: String(r[COL_LETTER - 1] || '').trim(),
        res: r.slice(COL_RES_FIRST - 1, COL_RES_LAST)
      });
    }
  }
  return { ok: true, disciplines: disciplines, data: data, updated: new Date().toISOString() };
}

function columnLetter(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
  return s;
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

// ===== КОНЕЦ ФАЙЛА. Если этой строки в редакторе Apps Script нет — код вставлен не полностью. =====
