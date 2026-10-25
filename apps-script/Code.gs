// ===== НАЧАЛО ФАЙЛА. Вставляйте весь файл целиком (около 640 строк), до метки «КОНЕЦ ФАЙЛА». =====
// Apps Script для таблицы «протокол главный».
// Развёртывание: Развернуть → Управление развёртываниями → ✏️ → Версия: «Новая версия» → Развернуть.
// Доступ: «Выполнять от имени: Я», «У кого есть доступ: Все».
// После вставки кода один раз выполните функцию authorize (кнопка «Выполнить») и дайте разрешения:
// скрипту нужен доступ к таблице «Регистрация ГТО» и к Google Диску (для архивных копий).

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

// Таблица «Регистрация ГТО»: лист на каждую дату, столбцы A:G — №, ФИО, УИН, Ступень, пол, класс, Буква
var REG_SPREADSHEET_ID = '1jjEFzNrZcsTDyyhd05oN8iIvo7scwmK0E57YhGEnNL8';
var REG_SKIP_SHEETS = ['основной'];          // шаблон, не дата
var CONTROL_SHEET = 'управление';            // сюда пишется, какая дата загружена
// Листы нормативов: при загрузке новой даты очищаются введённые результаты (номера в A и формулы остаются)
var RESULT_SHEETS = ['отжимание', 'подтягивание', 'прыжки', 'пресс', 'наклон', 'челнок 3х10',
                     '30 метров', 'длинная', 'метание', 'стрельба', '60м'];

function doGet(e) {
  var view = e && e.parameter && e.parameter.view;
  if (view === 'results') return jsonOut(getResults());
  if (view === 'dates') {
    try { return jsonOut(listDates()); } catch (err) { return jsonOut({ ok: false, error: String(err) }); }
  }
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonOut({ ok: true, data: [], event: currentEvent() });
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
  return jsonOut({ ok: true, data: data, event: currentEvent() });
}

function doPost(e) {
  // Одна запись за раз — несколько телефонов не мешают друг другу
  var lock = LockService.getScriptLock();
  try {
    var params = JSON.parse(e.postData.contents);
    lock.waitLock(params.action === 'loadDate' ? 30000 : 10000);
    if (params.action === 'add') return jsonOut(addParticipant(params));
    if (params.action === 'loadDate') return jsonOut(loadDate(params.sheet, params.mode));
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
  var norms = [];
  try { norms = getNorms(); } catch (e) {}
  return { ok: true, disciplines: disciplines, data: data, norms: norms, event: currentEvent(), updated: new Date().toISOString() };
}

function columnLetter(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
  return s;
}

// ===================== Загрузка участников из «Регистрации ГТО» =====================

// Список дат (листов) регистрации с числом участников
function listDates() {
  var reg = SpreadsheetApp.openById(REG_SPREADSHEET_ID);
  var dates = [];
  reg.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (REG_SKIP_SHEETS.indexOf(name.trim().toLowerCase()) !== -1) return;
    var count = readRegistration(sh).length;
    if (count) dates.push({ name: name, count: count });
  });
  return { ok: true, dates: dates, current: currentEvent() };
}

function readRegistration(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 7).getDisplayValues(); // A:G
  var people = [], seq = 0;
  vals.forEach(function (r) {
    var fio = String(r[1]).replace(/\s+/g, ' ').trim();
    if (!fio) return;
    seq++;
    var num = parseInt(r[0], 10);
    people.push({
      num: num > 0 ? num : seq,
      fio: fio,
      uin: String(r[2]).trim(),
      stupen: String(r[3]).replace(/\s+/g, ' ').trim(),
      pol: String(r[4]).trim().toLowerCase(),                // «Мужской» → «мужской», как в протоколе
      klass: parseInt(String(r[5]), 10) || String(r[5]).trim(), // «3 класс» → 3
      letter: String(r[6]).trim()
    });
  });
  return people;
}

// mode: 'replace' — архив + очистка + загрузка; 'append' — дописать тех, кого ещё нет (по УИН)
function loadDate(name, mode) {
  name = String(name || '');
  var reg = SpreadsheetApp.openById(REG_SPREADSHEET_ID);
  var src = reg.getSheetByName(name);
  if (!src) return { ok: false, error: 'в регистрации нет листа «' + name + '»' };
  var people = readRegistration(src);
  if (!people.length) return { ok: false, error: 'на листе «' + name + '» нет участников' };
  var sheet = getSheet();
  sortByFio(people);

  if (mode === 'append') {
    var res = appendPeople(sheet, people);
    setControl(name, null, null);
    return { ok: true, mode: 'append', added: res.added, skipped: res.skipped, event: name };
  }

  // Новая дата: алфавитный порядок и сквозные номера 1, 2, 3…
  people.forEach(function (p, i) { p.num = i + 1; });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archive = hasParticipants(sheet) ? archiveCopy(ss) : null;
  clearInputs(sheet, 2, 1, COL_LETTER);           // «Список»: A:S, формулы остаются
  ss.getSheets().forEach(function (sh) {
    if (RESULT_SHEETS.indexOf(normName(sh.getName())) === -1) return;
    var lastCol = sh.getLastColumn();
    if (lastCol >= 2) clearInputs(sh, 2, 2, lastCol); // столбец A с номерами не трогаем
  });
  writePeople(sheet, people, 2);
  try { ensureBadges(); } catch (e) {}
  setControl(name, people.length, archive);
  SpreadsheetApp.flush();
  return { ok: true, mode: 'replace', loaded: people.length, archive: archive, event: name };
}

// По алфавиту, без учёта регистра («РЯЗАНОВА» и «Рязанова» рядом), ё = е
function sortByFio(people) {
  var key = function (p) { return p.fio.toLowerCase().replace(/ё/g, 'е'); };
  people.sort(function (a, b) {
    var x = key(a), y = key(b);
    return x < y ? -1 : x > y ? 1 : 0;
  });
}

function normName(n) {
  return String(n).toLowerCase().replace(/x/g, 'х').replace(/\s+/g, ' ').trim();
}

function hasParticipants(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  return sheet.getRange(2, COL_FIO, last - 1, 1).getValues().some(function (r) { return String(r[0]).trim() !== ''; });
}

// Архив протокола перед загрузкой новой даты.
// Основной способ — полная копия файла в ту же папку Диска (нужно разрешение на Google Диск).
// Если такого разрешения нет — новая таблица в «Моём диске» с содержимым всех листов (как оно видно).
function archiveCopy(ss) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  var nm = ss.getName() + ' — архив «' + (currentEvent() || 'без даты') + '» (' + stamp + ')';
  try {
    var file = DriveApp.getFileById(ss.getId());
    var parents = file.getParents();
    var copy = parents.hasNext() ? file.makeCopy(nm, parents.next()) : file.makeCopy(nm);
    return { name: nm, url: copy.getUrl(), full: true };
  } catch (e) {
    return archiveValues(ss, nm);
  }
}

function archiveValues(ss, nm) {
  var arch = SpreadsheetApp.create(nm);
  var tmp = arch.getSheets()[0];
  tmp.setName('__tmp__' + new Date().getTime());
  ss.getSheets().forEach(function (sh, i) {
    var t = arch.insertSheet(sh.getName(), i);
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) return;
    if (t.getMaxColumns() < lc) t.insertColumnsAfter(t.getMaxColumns(), lc - t.getMaxColumns());
    if (t.getMaxRows() < lr) t.insertRowsAfter(t.getMaxRows(), lr - t.getMaxRows());
    var vals = sh.getRange(1, 1, lr, lc).getDisplayValues().map(function (row) {
      // текст, начинающийся с «=», не должен стать формулой
      return row.map(function (v) { return /^[=+]/.test(v) ? "'" + v : v; });
    });
    t.getRange(1, 1, lr, lc).setNumberFormat('@').setValues(vals);
  });
  arch.deleteSheet(tmp);
  return { name: nm, url: arch.getUrl(), full: false };
}

// Очищает введённые значения, не трогая формулы
function clearInputs(sheet, firstRow, firstCol, lastCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < firstRow || lastCol < firstCol) return;
  var n = lastRow - firstRow + 1, w = lastCol - firstCol + 1;
  var f = sheet.getRange(firstRow, firstCol, n, w).getFormulas();
  for (var c = 0; c < w; c++) {
    var start = -1;
    for (var r = 0; r <= n; r++) {
      var isFormula = r < n && f[r][c] !== '';
      if (r < n && !isFormula) { if (start < 0) start = r; continue; }
      if (start >= 0) {
        sheet.getRange(firstRow + start, firstCol + c, r - start, 1).clearContent();
        start = -1;
      }
    }
  }
}

function writePeople(sheet, people, startRow) {
  var n = people.length;
  var need = startRow + n - 1;
  if (need > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
  // Оформление и флажок «пришёл» — как у строки 2
  if (n > 0 && startRow + n - 1 > 2) {
    var lastCol = sheet.getLastColumn();
    var from = Math.max(startRow, 3);
    sheet.getRange(2, 1, 1, lastCol).copyTo(sheet.getRange(from, 1, startRow + n - from, lastCol), { formatOnly: true });
  }
  var dv = sheet.getRange(2, COL_PRESENT).getDataValidation();
  if (dv) sheet.getRange(startRow, COL_PRESENT, n, 1).setDataValidation(dv);
  sheet.getRange(startRow, COL_UIN, n, 1).setNumberFormat('@'); // УИН — текстом
  sheet.getRange(startRow, COL_NUM, n, 5).setValues(people.map(function (p) {
    return [p.num, p.fio, p.uin, p.stupen, p.pol];
  }));
  sheet.getRange(startRow, COL_PRESENT, n, 1).setValues(people.map(function () { return [false]; }));
  sheet.getRange(startRow, COL_CLASS, n, 2).setValues(people.map(function (p) { return [p.klass, p.letter]; }));
}

// Дописать новых (по УИН) после последнего ФИО; номер — из регистрации, если свободен, иначе следующий
function appendPeople(sheet, people) {
  var last = sheet.getLastRow();
  var have = {}, nums = {}, maxNum = 0, lastFioRow = 1;
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, COL_UIN).getValues().forEach(function (r, i) {
      if (String(r[COL_FIO - 1]).trim() === '') return;
      lastFioRow = i + 2;
      have[String(r[COL_UIN - 1]).replace(/\D/g, '')] = true;
      var n = parseInt(r[COL_NUM - 1], 10);
      if (n > 0) { nums[n] = true; if (n > maxNum) maxNum = n; }
    });
  }
  var add = [], skipped = 0;
  people.forEach(function (p) {
    var key = String(p.uin).replace(/\D/g, '');
    if (key && have[key]) { skipped++; return; }
    if (nums[p.num]) p.num = ++maxNum;
    nums[p.num] = true;
    if (p.num > maxNum) maxNum = p.num;
    have[key] = true;
    add.push(p);
  });
  if (add.length) writePeople(sheet, add, lastFioRow + 1);
  return { added: add.length, skipped: skipped };
}

// Что сейчас загружено — хранится в свойствах документа и видно на листе «управление»
function currentEvent() {
  return PropertiesService.getDocumentProperties().getProperty('gto_event') || '';
}

function setControl(name, count, archive) {
  PropertiesService.getDocumentProperties().setProperty('gto_event', name);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET);
  if (!sh) return;
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  sh.getRange(1, 1, 2, 2).setValues([['Загруженная дата', name], ['Обновлено', stamp]]);
  if (count !== null) sh.getRange(3, 1, 1, 2).setValues([['Участников', count]]);
  if (archive) sh.getRange(4, 1, 1, 2).setValues([['Архив предыдущей', '=HYPERLINK("' + archive.url + '","' + archive.name.replace(/"/g, "'") + '")']]);
}

// ===================== Подсветка знаков ГТО (золото / серебро / бронза) =====================
// Лист «Нормативы»: пороги по ступени, полу и столбцу протокола. Служебный скрытый лист «_знаки»
// считает уровень каждой ячейки результатов (0–3), условное форматирование «Списка» красит ячейку.

var NORMS_SHEET = 'Нормативы';
var BADGE_SHEET = '_знаки';
var BADGE_COLORS = { 3: '#FFD966', 2: '#D9D9D9', 1: '#F4B183' }; // золото, серебро, бронза
var NORMS_HEAD = ['Ступень', 'Пол', 'Столбец протокола', 'Испытание', 'Лучше', 'Формат',
                  'Бронза', 'Серебро', 'Золото', 'ключ (служебное)', 'бронза, число', 'серебро, число', 'золото, число'];

// III ступень (10–11 лет), таблица Мосгорспорта. [столбец, испытание М, испытание Ж, лучше, формат, М: б/с/з, Ж: б/с/з]
var NORMS_III = [
  ['30м', 'Бег 30 м (с)', 'Бег 30 м (с)', 'меньше', 'сек', [6.3, 5.9, 5.3], [6.5, 6.1, 5.5]],
  ['длинная', 'Бег на 1000 м (мин:с)', 'Бег на 1000 м (мин:с)', 'меньше', 'мин:сек', ['6:20', '5:41', '4:42'], ['6:40', '6:12', '5:03']],
  ['подтягивание', 'Подтягивание из виса на высокой перекладине', 'Подтягивание из виса лёжа на низкой перекладине 90 см', 'больше', 'число', [2, 4, 6], [7, 10, 16]],
  ['отжимание', 'Сгибание и разгибание рук в упоре лёжа на полу', 'Сгибание и разгибание рук в упоре лёжа на полу', 'больше', 'число', [10, 13, 22], [5, 7, 13]],
  ['наклон', 'Наклон вперёд стоя на скамье (см)', 'Наклон вперёд стоя на скамье (см)', 'больше', 'число', [2, 4, 8], [3, 5, 11]],
  ['челнок', 'Челночный бег 3×10 м (с)', 'Челночный бег 3×10 м (с)', 'меньше', 'сек', [9.7, 9.2, 8.4], [10.1, 9.4, 8.6]],
  ['прыжки', 'Прыжок в длину с места (см)', 'Прыжок в длину с места (см)', 'больше', 'число', [128, 142, 162], [118, 132, 152]],
  ['пресс', 'Поднимание туловища за 1 мин (раз)', 'Поднимание туловища за 1 мин (раз)', 'больше', 'число', [26, 33, 43], [23, 28, 37]],
  ['метание', 'Метание мяча 150 г (м)', 'Метание мяча 150 г (м)', 'больше', 'число', [18, 22, 28], [12, 15, 20]]
];

function normKeyPart(v) { return String(v == null ? '' : v).trim().toLowerCase(); }

// Число результата/порога в единицах сравнения. «мин:сек»: «6:20» → 380 с, 12,14 → 12 мин 14 с = 734 с
function normValue(v, fmt) {
  if (v === '' || v === null || v === undefined) return '';
  if (fmt === 'мин:сек') {
    if (v instanceof Date) return v.getHours() * 60 + v.getMinutes(); // «6:20», распознанное как время
    var t = String(v).trim();
    var mm = t.match(/^(\d+):(\d{1,2})/);
    if (mm) return +mm[1] * 60 + +mm[2];
    var x = parseFloat(t.replace(',', '.'));
    if (isNaN(x)) return '';
    var m = Math.floor(x);
    return m * 60 + Math.round((x - m) * 100);
  }
  var n = typeof v === 'number' ? v : parseFloat(String(v).trim().replace(',', '.'));
  return isNaN(n) ? '' : n;
}

// Заголовок «Списка», соответствующий столбцу норматива («длинная» может называться «длинная (1000 м)»)
function findListHeader(headers, name) {
  var want = normKeyPart(name);
  for (var i = 0; i < headers.length; i++) {
    var h = normKeyPart(headers[i]);
    if (h && (h === want || h.indexOf(want) === 0)) return String(headers[i]).trim();
  }
  return name;
}

function createNormsSheet(ss, list) {
  var sh = ss.insertSheet(NORMS_SHEET);
  var headers = list.getRange(1, COL_RES_FIRST, 1, COL_RES_LAST - COL_RES_FIRST + 1).getDisplayValues()[0];
  var rows = [];
  NORMS_III.forEach(function (n) {
    var col = findListHeader(headers, n[0]);
    rows.push(['III', 'мужской', col, n[1], n[3], n[4]].concat(n[5]));
    rows.push(['III', 'женский', col, n[2], n[3], n[4]].concat(n[6]));
  });
  sh.getRange(1, 1, 1, NORMS_HEAD.length).setValues([NORMS_HEAD]).setFontWeight('bold');
  sh.getRange(2, 7, Math.max(rows.length, 1) + 200, 3).setNumberFormat('@'); // пороги — как написаны («6:20», «5,3»)
  sh.getRange(2, 1, rows.length, 9).setValues(rows.map(function (r) {
    return r.map(function (v) { return typeof v === 'number' ? String(v).replace('.', ',') : v; });
  }));
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 9);
  return sh;
}

// Служебные столбцы J:M — ключ и пороги числами (пересчитываются при правке листа)
function refreshNormNumbers(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, 9).getValues();
  var out = vals.map(function (r) {
    if (String(r[0]).trim() === '') return ['', '', '', ''];
    var fmt = normKeyPart(r[5]);
    return [
      String(r[0]).trim().toUpperCase() + '|' + normKeyPart(r[1]) + '|' + normKeyPart(r[2]),
      normValue(r[6], fmt), normValue(r[7], fmt), normValue(r[8], fmt)
    ];
  });
  sh.getRange(2, 10, out.length, 4).setValues(out);
  sh.getRange(1, 10, 1, 4).setValues([NORMS_HEAD.slice(9)]).setFontColor('#999999');
}

// Служебный лист «_знаки» — четыре блока формул на каждую ячейку результатов «Списка» (F:P):
//   F:P   — уровень 0–3 (его читает условное форматирование)
//   V:AF  — номер строки норматива на листе «Нормативы»
//   AH:AR — результат числом (как в ячейке)
//   AT:BD — результат в единицах сравнения (для «мин:сек» — секунды)
// Используются только общие функции (INDEX, MATCH, VALUE…), без LET и регулярных выражений.
var BADGE_BLOCK = { level: 6, match: 22, num: 34, val: 46 };

function badgeFormulas(r, c) {
  var L = "'" + SHEET_NAME + "'!", N = "'" + NORMS_SHEET + "'!";
  var off = c - COL_RES_FIRST;
  var col = columnLetter(c);
  var mC = columnLetter(BADGE_BLOCK.match + off) + r;
  var xC = columnLetter(BADGE_BLOCK.num + off) + r;
  var vC = columnLetter(BADGE_BLOCK.val + off) + r;
  var raw = L + col + r;
  var idx = function (nc) { return 'INDEX(' + N + '$' + nc + ':$' + nc + ',' + mC + ')'; };
  var stupen = 'UPPER(LEFT(TRIM(' + L + '$D' + r + '),FIND(" ",TRIM(' + L + '$D' + r + ')&" ")-1))';
  var key = stupen + '&"|"&LOWER(TRIM(' + L + '$E' + r + '))&"|"&LOWER(TRIM(' + L + col + '$1))';
  var match = '=IF(OR(' + raw + '="",' + L + '$B' + r + '=""),"",IFERROR(MATCH(' + key + ',' + N + '$J:$J,0),""))';
  var num = '=IF(' + mC + '="","",IF(ISNUMBER(' + raw + '),' + raw + ',IFERROR(VALUE(' + raw + '),IFERROR(VALUE(SUBSTITUTE(' + raw + ',",",".")),IFERROR(VALUE(SUBSTITUTE(' + raw + ',".",",")),"")))))';
  var val = '=IF(' + xC + '="","",IF(' + idx('F') + '="мин:сек",INT(' + xC + ')*60+ROUND((' + xC + '-INT(' + xC + '))*100,0),' + xC + '))';
  var b = idx('K'), sv = idx('L'), g = idx('M');
  var cmp = function (t, op) { return 'AND(' + t + '<>"",' + vC + op + t + ')'; };
  var level = '=IF(' + vC + '="","",IF(' + idx('E') + '="больше",' +
    'IF(' + cmp(g, '>=') + ',3,IF(' + cmp(sv, '>=') + ',2,IF(' + cmp(b, '>=') + ',1,0))),' +
    'IF(' + cmp(g, '<=') + ',3,IF(' + cmp(sv, '<=') + ',2,IF(' + cmp(b, '<=') + ',1,0)))))';
  return { level: level, match: match, num: num, val: val };
}

function setupBadges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = getSheet();
  var norms = ss.getSheetByName(NORMS_SHEET) || createNormsSheet(ss, list);
  refreshNormNumbers(norms);

  var rows = list.getMaxRows();
  var width = COL_RES_LAST - COL_RES_FIRST + 1;
  var needCols = BADGE_BLOCK.val + width - 1;
  var helper = ss.getSheetByName(BADGE_SHEET) || ss.insertSheet(BADGE_SHEET);
  if (helper.getMaxRows() < rows) helper.insertRowsAfter(helper.getMaxRows(), rows - helper.getMaxRows());
  if (helper.getMaxColumns() < needCols) helper.insertColumnsAfter(helper.getMaxColumns(), needCols - helper.getMaxColumns());
  helper.getRange(1, 1).setValue('Служебный лист: уровень знака для ячеек «Списка» (3 — золото, 2 — серебро, 1 — бронза). Не редактировать.');
  var blocks = { level: [], match: [], num: [], val: [] };
  for (var r = 2; r <= rows; r++) {
    var rowF = { level: [], match: [], num: [], val: [] };
    for (var c = COL_RES_FIRST; c <= COL_RES_LAST; c++) {
      var f = badgeFormulas(r, c);
      rowF.level.push(f.level); rowF.match.push(f.match); rowF.num.push(f.num); rowF.val.push(f.val);
    }
    Object.keys(blocks).forEach(function (k) { blocks[k].push(rowF[k]); });
  }
  Object.keys(blocks).forEach(function (k) {
    helper.getRange(2, BADGE_BLOCK[k], rows - 1, width).setFormulas(blocks[k]);
  });
  helper.hideSheet();

  // Условное форматирование: наши правила — первыми, чужие правила сохраняем
  var range = list.getRange(2, COL_RES_FIRST, rows - 1, width);
  var others = list.getConditionalFormatRules().filter(function (rule) {
    var bc = rule.getBooleanCondition();
    var vals = bc ? bc.getCriteriaValues() : [];
    return !(vals.length && String(vals[0]).indexOf(BADGE_SHEET) !== -1);
  });
  var ours = [3, 2, 1].map(function (lvl) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=INDIRECT("\'' + BADGE_SHEET + '\'!"&ADDRESS(ROW(),COLUMN()))=' + lvl)
      .setBackground(BADGE_COLORS[lvl])
      .setRanges([range])
      .build();
  });
  list.setConditionalFormatRules(ours.concat(others));
  return { ok: true, rows: rows - 1 };
}

// Если «Список» вырос (новая дата с большим числом участников) — достраиваем подсветку
function ensureBadges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var helper = ss.getSheetByName(BADGE_SHEET);
  if (helper && helper.getMaxRows() < getSheet().getMaxRows()) setupBadges();
}

// Правка листа «Нормативы» — сразу пересчитываем служебные пороги
function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === NORMS_SHEET) refreshNormNumbers(e.range.getSheet());
  } catch (err) {}
}

// Нормативы для сайта
function getNorms() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NORMS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues()
    .filter(function (r) { return String(r[9]).trim() !== ''; })
    .map(function (r) {
      return { key: String(r[9]), dir: normKeyPart(r[4]), fmt: normKeyPart(r[5]), b: r[10], s: r[11], g: r[12] };
    });
}

// ===================== Меню «ГТО» в самой таблице =====================

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (/архив/i.test(ss.getName())) return; // в архивных копиях загрузка не нужна
  SpreadsheetApp.getUi().createMenu('ГТО')
    .addItem('Загрузить дату из регистрации…', 'showLoadDialog')
    .addItem('Подсветка знаков: создать / обновить', 'menuSetupBadges')
    .addToUi();
}

function menuSetupBadges() {
  var res = setupBadges();
  SpreadsheetApp.getUi().alert('Подсветка знаков настроена: строк «Списка» — ' + res.rows +
    '.\nПороги — на листе «' + NORMS_SHEET + '», их можно править.');
}

function showLoadDialog() {
  var info = listDates();
  var esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); };
  var opts = info.dates.map(function (d) {
    return '<option value="' + esc(d.name) + '"' + (d.name === info.current ? ' selected' : '') + '>' +
      esc(d.name) + ' — ' + d.count + ' чел.</option>';
  }).join('');
  var html =
    '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p>Сейчас загружено: <b>' + esc(info.current || '—') + '</b></p>' +
    '<select id="d" style="width:100%;padding:6px;font-size:14px">' + opts + '</select>' +
    '<p style="color:#666;font-size:12px">«Загрузить» — копия протокола уйдёт в архив на Google Диск, ' +
    'затем «Список», явка и результаты на листах нормативов очистятся и загрузятся участники выбранной даты.<br>' +
    '«Догрузить новых» — добавит только тех, кого ещё нет в «Списке» (по УИН), ничего не очищая.</p>' +
    '<button onclick="go(\'replace\')" style="padding:8px 12px">Загрузить</button> ' +
    '<button onclick="go(\'append\')" style="padding:8px 12px">Догрузить новых</button>' +
    '<p id="m" style="margin-top:12px"></p></div>' +
    '<script>function go(mode){var d=document.getElementById("d").value,m=document.getElementById("m");' +
    'if(mode==="replace"&&!confirm("Загрузить «"+d+"»?\\nТекущий протокол уйдёт в архив, список и результаты очистятся."))return;' +
    'm.textContent="Загружаем…";document.querySelectorAll("button").forEach(function(b){b.disabled=true});' +
    'google.script.run.withSuccessHandler(function(r){m.textContent=r.ok?(r.mode==="append"?"Добавлено: "+r.added+", уже были: "+r.skipped:"Загружено участников: "+r.loaded+(r.archive?". Архив: "+r.archive.name:"")):("Ошибка: "+r.error);' +
    'document.querySelectorAll("button").forEach(function(b){b.disabled=false})})' +
    '.withFailureHandler(function(e){m.textContent="Ошибка: "+e.message;document.querySelectorAll("button").forEach(function(b){b.disabled=false})})' +
    '.menuLoad(d,mode)}</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(330), 'Загрузка даты из регистрации');
}

function menuLoad(name, mode) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return loadDate(name, mode); } finally { lock.releaseLock(); }
}

// Запустите один раз из редактора, чтобы выдать скрипту разрешения
function authorize() {
  SpreadsheetApp.openById(REG_SPREADSHEET_ID).getName();
  DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getName();
  PropertiesService.getDocumentProperties().getProperty('gto_event');
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
