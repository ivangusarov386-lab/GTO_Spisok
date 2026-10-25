// ===== НАЧАЛО ФАЙЛА. Вставляйте весь файл целиком (около 760 строк), до метки «КОНЕЦ ФАЙЛА». =====
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
    if (params.action === 'renumber') return jsonOut(sortAndRenumber());
    if (params.action && params.action !== 'mark') {
      return jsonOut({ ok: false, error: 'неизвестное действие: ' + params.action });
    }
    var row = parseInt(params.row, 10);
    var present = !!params.present;
    var sheet = getSheet();

    // Строки могли сдвинуть (сортировка, перенумерация) — сверяем УИН, а если его нет — номер,
    // и при расхождении ищем строку участника заново.
    var uin = String(params.uin == null ? '' : params.uin).trim();
    if (uin) {
      var hereUin = row >= 2 ? String(sheet.getRange(row, COL_UIN).getValue()).trim() : '';
      if (hereUin !== uin) {
        row = findRowByUin(sheet, uin);
        if (!row) return jsonOut({ ok: false, error: 'участник с УИН ' + uin + ' не найден в листе' });
      }
    } else if (params.num !== undefined && params.num !== null && params.num !== '') {
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
  SpreadsheetApp.flush();
  try { paintBadges(); } catch (e) {}
  setControl(name, people.length, archive);
  SpreadsheetApp.flush();
  return { ok: true, mode: 'replace', loaded: people.length, archive: archive, event: name };
}

// ===================== Упорядочить «Список» по ФИО и перенумеровать 1…N =====================
// Строки листов нормативов связаны со строками «Списка» (строка 5 = участник в строке 5 «Списка»),
// поэтому вместе с участниками переставляются и введённые результаты. Формулы не трогаем.
function permuteRows(sh, firstCol, lastCol, n, order) {
  if (lastCol < firstCol || n < 1) return;
  var w = lastCol - firstCol + 1;
  var rg = sh.getRange(2, firstCol, n, w);
  var vals = rg.getValues(), forms = rg.getFormulas();
  for (var c = 0; c < w; c++) {
    var hasF = false, allF = true;
    for (var r = 0; r < n; r++) { if (forms[r][c]) hasF = true; else allF = false; }
    if (hasF) continue;                                         // столбец с формулами остаётся как есть
    var col = vals.map(function (row) { return row[c]; });
    sh.getRange(2, firstCol + c, n, 1).setValues(order.map(function (i) { return [col[i]]; }));
  }
}

function sortAndRenumber() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = getSheet();
  var last = list.getLastRow();
  if (last < 2) return { ok: true, count: 0, moved: 0 };
  var fio = list.getRange(2, COL_FIO, last - 1, 1).getValues();
  var n = fio.length;
  while (n > 0 && String(fio[n - 1][0]).trim() === '') n--;
  if (!n) return { ok: true, count: 0, moved: 0 };
  var key = function (i) { return String(fio[i][0]).trim().toLowerCase().replace(/ё/g, 'е'); };
  var order = [];
  for (var i = 0; i < n; i++) order.push(i);
  order.sort(function (a, b) { var x = key(a), y = key(b); return x < y ? -1 : x > y ? 1 : a - b; });
  var moved = order.filter(function (v, i) { return v !== i; }).length;

  if (moved) {
    permuteRows(list, 1, COL_LETTER, n, order);
    ss.getSheets().forEach(function (sh) {
      if (RESULT_SHEETS.indexOf(normName(sh.getName())) === -1) return;
      permuteRows(sh, 2, sh.getLastColumn(), n, order);
    });
  }
  var nums = [];
  for (var k = 1; k <= n; k++) nums.push([k]);
  list.getRange(2, COL_NUM, n, 1).setValues(nums);
  SpreadsheetApp.flush();
  try { paintBadges(); } catch (e) {}
  return { ok: true, count: n, moved: moved };
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
// Лист «Нормативы»: пороги по ступени, полу и столбцу протокола. Скрипт сам красит ячейки результатов
// «Списка» при каждой правке таблицы (paintBadges).

var NORMS_SHEET = 'Нормативы';
var BADGE_SHEET = '_знаки'; // служебный лист прежней версии — удаляется при настройке
var BADGE_COLORS = { 3: '#FFD966', 2: '#D9D9D9', 1: '#F4B183' }; // золото, серебро, бронза
var NORMS_HEAD = ['Ступень', 'Пол', 'Столбец протокола', 'Испытание', 'Лучше', 'Формат',
                  'Бронза', 'Серебро', 'Золото', 'ключ (служебное)', 'бронза, число', 'серебро, число', 'золото, число'];

// Нормативы Мосгорспорта по ступеням. [столбец, испытание М, испытание Ж, лучше, формат, М: б/с/з, Ж: б/с/з]
// null — у этого пола такого испытания нет. Подтягивание: юноши/мальчики — высокая перекладина, девушки/девочки — низкая.
// Стрельба — пневматическая винтовка с диоптрическим прицелом или «электронное оружие».
var PULL_M = 'Подтягивание из виса на высокой перекладине';
var PULL_W = 'Подтягивание из виса лёжа на низкой перекладине 90 см';
var PUSH = 'Сгибание и разгибание рук в упоре лёжа на полу';
var BEND = 'Наклон вперёд стоя на скамье (см)';
var SHUTTLE = 'Челночный бег 3×10 м (с)';
var JUMP = 'Прыжок в длину с места (см)';
var PRESS = 'Поднимание туловища за 1 мин (раз)';
var SHOOT = 'Стрельба: диоптрический прицел / электронное оружие (очки)';
var TENNIS = 'Метание теннисного мяча в цель (попаданий)';
var NORMS = {
  'I': [ // 6–7 лет
    ['30м', 'Бег 30 м (с)', 'Бег 30 м (с)', 'меньше', 'сек', [9.6, 8.6, 8.3], [10.1, 9.1, 8.8]],
    ['длинная', 'Шестиминутный бег (м)', 'Шестиминутный бег (м)', 'больше', 'число', [400, 600, 800], [300, 500, 700]],
    ['наклон', BEND, BEND, 'больше', 'число', [1, 3, 7], [3, 5, 9]],
    ['челнок', SHUTTLE, SHUTTLE, 'меньше', 'сек', [11.2, 10.8, 9.9], [11.7, 11.3, 10.8]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [85, 95, 115], [77, 87, 107]],
    ['пресс', 'Поднимание туловища за 30 с (раз)', 'Поднимание туловища за 30 с (раз)', 'больше', 'число', [11, 12, 15], [10, 11, 14]],
    ['метание', TENNIS, TENNIS, 'больше', 'число', [2, 3, 4], [1, 2, 3]]
  ],
  'II': [ // 8–9 лет
    ['30м', 'Бег 30 м (с)', 'Бег 30 м (с)', 'меньше', 'сек', [7.1, 6.5, 5.9], [7.3, 6.6, 6.1]],
    ['длинная', 'Смешанное передвижение 1000 м (мин:с)', 'Смешанное передвижение 1000 м (мин:с)', 'меньше', 'мин:сек', ['7:21', '6:30', '5:12'], ['7:45', '6:55', '5:55']],
    ['подтягивание', PULL_M, PULL_W, 'больше', 'число', [2, 4, 5], [4, 7, 12]],
    ['отжимание', PUSH, PUSH, 'больше', 'число', [5, 11, 18], [4, 7, 12]],
    ['наклон', BEND, BEND, 'больше', 'число', [1, 3, 7], [3, 5, 9]],
    ['челнок', SHUTTLE, SHUTTLE, 'меньше', 'сек', [10.5, 9.8, 9.1], [10.8, 10.2, 9.4]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [108, 122, 142], [103, 117, 137]],
    ['пресс', PRESS, PRESS, 'больше', 'число', [20, 25, 37], [17, 22, 32]],
    ['метание', TENNIS, TENNIS, 'больше', 'число', [2, 3, 4], [1, 2, 3]]
  ],
  'III': [ // 10–11 лет
    ['30м', 'Бег 30 м (с)', 'Бег 30 м (с)', 'меньше', 'сек', [6.3, 5.9, 5.3], [6.5, 6.1, 5.5]],
    ['длинная', 'Бег на 1000 м (мин:с)', 'Бег на 1000 м (мин:с)', 'меньше', 'мин:сек', ['6:20', '5:41', '4:42'], ['6:40', '6:12', '5:03']],
    ['подтягивание', PULL_M, PULL_W, 'больше', 'число', [2, 4, 6], [7, 10, 16]],
    ['отжимание', PUSH, PUSH, 'больше', 'число', [10, 13, 22], [5, 7, 13]],
    ['наклон', BEND, BEND, 'больше', 'число', [2, 4, 8], [3, 5, 11]],
    ['челнок', SHUTTLE, SHUTTLE, 'меньше', 'сек', [9.7, 9.2, 8.4], [10.1, 9.4, 8.6]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [128, 142, 162], [118, 132, 152]],
    ['пресс', PRESS, PRESS, 'больше', 'число', [26, 33, 43], [23, 28, 37]],
    ['метание', 'Метание мяча 150 г (м)', 'Метание мяча 150 г (м)', 'больше', 'число', [18, 22, 28], [12, 15, 20]]
  ],
  'IV': [ // 12–13 лет
    ['30м', 'Бег 30 м (с)', 'Бег 30 м (с)', 'меньше', 'сек', [5.8, 5.4, 5.0], [6.1, 5.7, 5.2]],
    ['60м', 'Бег 60 м (с)', 'Бег 60 м (с)', 'меньше', 'сек', [11.1, 10.2, 9.4], [11.5, 10.7, 9.9]],
    ['длинная', 'Бег на 1500 м (мин:с)', 'Бег на 1500 м (мин:с)', 'меньше', 'мин:сек', ['8:30', '7:55', '6:40'], ['9:10', '8:15', '7:03']],
    ['подтягивание', PULL_M, PULL_W, 'больше', 'число', [3, 5, 8], [8, 12, 18]],
    ['отжимание', PUSH, PUSH, 'больше', 'число', [12, 18, 29], [6, 9, 15]],
    ['наклон', BEND, BEND, 'больше', 'число', [3, 5, 9], [4, 6, 13]],
    ['челнок', SHUTTLE, SHUTTLE, 'меньше', 'сек', [9.2, 8.5, 7.7], [9.6, 8.9, 8.0]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [147, 162, 183], [132, 147, 167]],
    ['пресс', PRESS, PRESS, 'больше', 'число', [31, 37, 47], [27, 31, 41]],
    ['метание', 'Метание мяча 150 г (м)', 'Метание мяча 150 г (м)', 'больше', 'число', [23, 27, 34], [15, 19, 23]],
    ['стрельба', SHOOT, SHOOT, 'больше', 'число', [10, 20, 25], [10, 20, 25]]
  ],
  'V': [ // 14–15 лет
    ['30м', 'Бег 30 м (с)', 'Бег 30 м (с)', 'меньше', 'сек', [5.4, 5.0, 4.6], [5.7, 5.3, 4.9]],
    ['60м', 'Бег 60 м (с)', 'Бег 60 м (с)', 'меньше', 'сек', [9.7, 9.1, 8.1], [10.8, 10.2, 9.5]],
    ['длинная', 'Бег на 2000 м (мин:с)', 'Бег на 2000 м (мин:с)', 'меньше', 'мин:сек', ['10:10', '9:27', '8:00'], ['12:40', '11:27', '9:55']],
    ['подтягивание', PULL_M, PULL_W, 'больше', 'число', [5, 9, 13], [9, 13, 19]],
    ['отжимание', PUSH, PUSH, 'больше', 'число', [19, 25, 37], [7, 11, 16]],
    ['наклон', BEND, BEND, 'больше', 'число', [4, 6, 11], [5, 8, 15]],
    ['челнок', SHUTTLE, SHUTTLE, 'меньше', 'сек', [8.2, 7.7, 7.1], [9.1, 8.7, 7.9]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [167, 193, 218], [148, 162, 183]],
    ['пресс', PRESS, PRESS, 'больше', 'число', [34, 40, 50], [31, 35, 44]],
    ['метание', 'Метание мяча 150 г (м)', 'Метание мяча 150 г (м)', 'больше', 'число', [30, 35, 41], [19, 21, 27]],
    ['стрельба', SHOOT, SHOOT, 'больше', 'число', [18, 25, 30], [18, 25, 30]]
  ],
  'VI': [ // 16–17 лет
    ['60м', 'Бег 60 м (с)', 'Бег 60 м (с)', 'меньше', 'сек', [9.0, 8.4, 7.9], [10.7, 9.9, 9.2]],
    ['длинная', 'Бег на 3000 м (мин:с)', 'Бег на 2000 м (мин:с)', 'меньше', 'мин:сек', ['15:20', '14:10', '12:20'], ['12:25', '11:10', '9:45']],
    ['подтягивание', PULL_M, PULL_W, 'больше', 'число', [8, 12, 15], [10, 14, 20]],
    ['отжимание', PUSH, PUSH, 'больше', 'число', [25, 32, 43], [8, 12, 17]],
    ['наклон', BEND, BEND, 'больше', 'число', [6, 8, 13], [7, 9, 16]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [192, 213, 235], [157, 173, 188]],
    ['пресс', PRESS, PRESS, 'больше', 'число', [35, 41, 51], [32, 37, 45]],
    ['метание', 'Метание снаряда 700 г (м)', 'Метание снаряда 500 г (м)', 'больше', 'число', [27, 30, 36], [12, 17, 22]],
    ['стрельба', SHOOT, SHOOT, 'больше', 'число', [18, 25, 30], [18, 25, 30]]
  ],
  'VII': [ // 18–19 лет
    ['60м', 'Бег 60 м (с)', 'Бег 60 м (с)', 'меньше', 'сек', [8.9, 8.4, 7.9], [10.7, 9.9, 9.2]],
    ['длинная', 'Бег на 3000 м (мин:с)', 'Бег на 2000 м (мин:с)', 'меньше', 'мин:сек', ['15:20', '14:10', '12:20'], ['12:20', '11:05', '9:40']],
    ['подтягивание', PULL_M, PULL_W, 'больше', 'число', [8, 12, 15], [10, 14, 20]],
    ['отжимание', PUSH, PUSH, 'больше', 'число', [25, 32, 43], [8, 12, 17]],
    ['наклон', BEND, BEND, 'больше', 'число', [6, 8, 13], [7, 9, 16]],
    ['прыжки', JUMP, JUMP, 'больше', 'число', [192, 213, 233], [157, 173, 188]],
    ['пресс', PRESS, PRESS, 'больше', 'число', [34, 41, 51], [31, 37, 45]],
    ['метание', 'Метание снаряда 700 г (м)', 'Метание снаряда 500 г (м)', 'больше', 'число', [27, 29, 36], [13, 16, 20]],
    ['стрельба', SHOOT, SHOOT, 'больше', 'число', [18, 25, 30], [18, 25, 30]]
  ]
};

function normKeyPart(v) { return String(v == null ? '' : v).trim().toLowerCase(); }

// «00:01:07» и «0:12:40» — ч:мм:сс; «12:40:00» — это «12:40», которое Google превратил во время
function hmsSeconds(h, m, sec) { return h > 0 && sec === 0 ? h * 60 + m : (h * 60 + m) * 60 + sec; }

// Число результата/порога в единицах сравнения. «мин:сек»: «6:20» → 380 с, 12,14 → 12 мин 14 с = 734 с
function normValue(v, fmt) {
  if (v === '' || v === null || v === undefined) return '';
  if (fmt === 'мин:сек') {
    if (v instanceof Date) return v.getHours() * 60 + v.getMinutes(); // «6:20», распознанное как время
    var t = String(v).trim();
    var hms = t.match(/^(\d+):(\d{1,2}):(\d{1,2})$/);                  // «00:01:07» → 1 мин 7 с
    if (hms) return hmsSeconds(+hms[1], +hms[2], +hms[3]);
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

function createNormsSheet(ss) {
  var sh = ss.insertSheet(NORMS_SHEET);
  sh.getRange(1, 1, 1, NORMS_HEAD.length).setValues([NORMS_HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// Дописывает на лист «Нормативы» строки, которых там ещё нет (ступень + пол + столбец).
// Уже существующие строки, в том числе исправленные вручную, не трогает.
function addMissingNorms(sh, list) {
  var headers = list.getRange(1, COL_RES_FIRST, 1, COL_RES_LAST - COL_RES_FIRST + 1).getDisplayValues()[0];
  var last = sh.getLastRow();
  var have = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
      have[String(r[0]).trim().toUpperCase() + '|' + normKeyPart(r[1]) + '|' + normKeyPart(r[2])] = true;
    });
  }
  var text = function (v) { return typeof v === 'number' ? String(v).replace('.', ',') : v; };
  var rows = [];
  Object.keys(NORMS).forEach(function (st) {
    NORMS[st].forEach(function (n) {
      var col = findListHeader(headers, n[0]);
      [['мужской', n[1], n[5]], ['женский', n[2], n[6]]].forEach(function (p) {
        if (!p[2] || have[st + '|' + p[0] + '|' + normKeyPart(col)]) return;
        rows.push([st, p[0], col, p[1], n[3], n[4]].concat(p[2].map(text)));
      });
    });
  });
  if (!rows.length) return 0;
  var start = Math.max(last, 1) + 1;
  sh.getRange(start, 7, rows.length, 3).setNumberFormat('@'); // пороги — как написаны («6:20», «5,3»)
  sh.getRange(start, 1, rows.length, 9).setValues(rows);
  sh.autoResizeColumns(1, 9);
  return rows.length;
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

// Цвет ставит сам скрипт (без формул и условного форматирования — они ломаются при русской локали таблицы).
// Перекрашивает только ячейки результатов F:P; чужую заливку не трогает, снимает лишь свои три цвета.
function paintBadges() {
  var list = getSheet();
  var map = {};
  getNorms().forEach(function (n) { map[n.key] = n; });
  var last = list.getLastRow();
  var fio = last >= 2 ? list.getRange(2, COL_FIO, last - 1, 1).getValues() : [];
  var people = fio.length;
  while (people > 0 && String(fio[people - 1][0]).trim() === '') people--;
  // строки, покрашенные в прошлый раз, тоже проходим — после загрузки короткого списка снимаем старые цвета
  var props = PropertiesService.getDocumentProperties();
  var n = Math.min(Math.max(people, +props.getProperty('badge_rows') || 0), list.getMaxRows() - 1);
  if (String(people) !== props.getProperty('badge_rows')) props.setProperty('badge_rows', String(people));
  if (n < 1) return 0;
  var width = COL_RES_LAST - COL_RES_FIRST + 1;
  var head = list.getRange(1, COL_RES_FIRST, 1, width).getDisplayValues()[0].map(normKeyPart);
  var info = list.getRange(2, 1, n, 5).getValues();             // D — ступень, E — пол
  var vals = list.getRange(2, COL_RES_FIRST, n, width).getValues();
  var range = list.getRange(2, COL_RES_FIRST, n, width);
  var bgs = range.getBackgrounds();
  var ours = {};
  Object.keys(BADGE_COLORS).forEach(function (k) { ours[BADGE_COLORS[k].toLowerCase()] = true; });
  var num = function (t) { return typeof t === 'number' ? t : null; };
  var changed = 0;
  for (var r = 0; r < n; r++) {
    if (r >= people) info[r] = ['', '', '', '', ''];            // участника уже нет — только снять свои цвета
    var st = (String(info[r][3]).trim().split(' ')[0] || '').toUpperCase();
    var prefix = st + '|' + normKeyPart(info[r][4]) + '|';
    for (var c = 0; c < width; c++) {
      var lvl = 0, nr = map[prefix + head[c]];
      if (nr && vals[r][c] !== '') {
        var v = normValue(vals[r][c], nr.fmt);
        if (v !== '') {
          var up = nr.dir === 'больше';
          var ok = function (t) { t = num(t); return t !== null && (up ? v >= t : v <= t); };
          lvl = ok(nr.g) ? 3 : ok(nr.s) ? 2 : ok(nr.b) ? 1 : 0;
        }
      }
      var cur = String(bgs[r][c]).toLowerCase();
      var want = lvl ? BADGE_COLORS[lvl].toLowerCase() : (ours[cur] ? '#ffffff' : cur);
      if (want !== cur) { bgs[r][c] = want; changed++; }
    }
  }
  if (changed) range.setBackgrounds(bgs);
  return changed;
}

function setupBadges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = getSheet();
  var norms = ss.getSheetByName(NORMS_SHEET) || createNormsSheet(ss);
  var added = addMissingNorms(norms, list);
  refreshNormNumbers(norms);

  // Убираем прежний вариант на формулах: служебный лист «_знаки» и его правила условного форматирования
  var helper = ss.getSheetByName(BADGE_SHEET);
  if (helper) ss.deleteSheet(helper);
  var rules = list.getConditionalFormatRules();
  var keep = rules.filter(function (rule) {
    var bc = rule.getBooleanCondition();
    var vals = bc ? bc.getCriteriaValues() : [];
    return !(vals.length && String(vals[0]).indexOf(BADGE_SHEET) !== -1);
  });
  if (keep.length !== rules.length) list.setConditionalFormatRules(keep);

  setupLongSheet();
  SpreadsheetApp.flush();
  paintBadges();
  return { ok: true, added: added };
}

// ===================== Длинный бег: результат в виде «мин:сек» =====================
// На листе «длинная» результат вводят в столбец B (текстом): «12,40», «12.4», «12:40» → «12:40».
// Целые числа не трогаем (на I ступени там метры шестиминутного бега). Столбец C просто повторяет B.
var LONG_SHEET = 'длинная';

function toMinSecText(v) {
  if (v instanceof Date) return v.getHours() + ':' + ('0' + v.getMinutes()).slice(-2); // «12:40», распознанное как время
  var t = String(v == null ? '' : v).trim();
  var hms = t.match(/^(\d+):(\d{1,2}):(\d{1,2})$/);                  // «00:01:07» → «1:07»
  if (hms) { var tot = hmsSeconds(+hms[1], +hms[2], +hms[3]); return Math.floor(tot / 60) + ':' + ('0' + tot % 60).slice(-2); }
  var m = t.match(/^(\d+)\s*[:.,]\s*(\d{1,2})$/);
  if (!m) return null;
  var sec = m[2].length === 1 ? +m[2] * 10 : +m[2];                                   // 12,4 → 12:40
  if (sec > 59) return null;
  return +m[1] + ':' + ('0' + sec).slice(-2);
}

function fixLongCells(range) {
  var vals = range.getValues(), forms = range.getFormulas(), changed = false;
  var out = vals.map(function (row, i) {
    return row.map(function (v, j) {
      if (forms[i][j]) return v;
      var t = toMinSecText(v);
      if (t === null || t === v) return v;
      changed = true;
      return t;
    });
  });
  if (changed) range.setNumberFormat('@').setValues(out);
}

// Настройка листа «длинная»: C = B (без пересчёта в «12,40»), уже введённые результаты — в «мин:сек»
function setupLongSheet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function (x) {
    return normName(x.getName()) === LONG_SHEET;
  })[0];
  if (!sh || sh.getMaxRows() < 2) return;
  var n = sh.getMaxRows() - 1;
  var f = [];
  for (var r = 2; r <= n + 1; r++) f.push(['=B' + r + '&""']);
  sh.getRange(2, 3, n, 1).setFormulas(f);
  var last = sh.getLastRow();
  if (last >= 2) fixLongCells(sh.getRange(2, 2, last - 1, 1));
}

// Любая правка (результат на листе норматива, порог на «Нормативах») — перекрашиваем «Список»
function onEdit(e) {
  try {
    var name = e && e.range ? e.range.getSheet().getName() : '';
    if (name === NORMS_SHEET) refreshNormNumbers(e.range.getSheet());
    if (normName(name) === LONG_SHEET && e.range.getColumn() <= 2 && e.range.getLastColumn() >= 2) {
      var rg = e.range, r0 = Math.max(rg.getRow(), 2), r1 = rg.getLastRow();
      if (r1 >= r0) fixLongCells(rg.getSheet().getRange(r0, 2, r1 - r0 + 1, 1));
    }
    paintBadges();
  } catch (err) {}
}

// Нормативы для сайта
function getNorms() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NORMS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues()
    .filter(function (r) { return String(r[9]).trim() !== ''; })
    .map(function (r) {
      return { key: String(r[9]), name: String(r[3]), dir: normKeyPart(r[4]), fmt: normKeyPart(r[5]), b: r[10], s: r[11], g: r[12] };
    });
}

// ===================== Меню «ГТО» в самой таблице =====================

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (/архив/i.test(ss.getName())) return; // в архивных копиях загрузка не нужна
  SpreadsheetApp.getUi().createMenu('ГТО')
    .addItem('Загрузить дату из регистрации…', 'showLoadDialog')
    .addItem('Подсветка знаков: создать / обновить', 'menuSetupBadges')
    .addItem('Упорядочить по алфавиту и перенумеровать', 'menuRenumber')
    .addToUi();
}

function menuRenumber() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Упорядочить по алфавиту',
    '«Список» отсортируется по ФИО, номера станут 1, 2, 3… по порядку. Явка и результаты переедут вместе с участниками.\n' +
    'Уже напечатанные стикеры с номерами станут неверными. Продолжить?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var res = sortAndRenumber();
  ui.alert(res.ok ? 'Готово: участников — ' + res.count + (res.moved ? ', поменяли место — ' + res.moved : ', порядок уже был верный') + '.' : 'Ошибка: ' + res.error);
}

function menuSetupBadges() {
  var res = setupBadges();
  SpreadsheetApp.getUi().alert('Подсветка знаков обновлена.' +
    (res.added ? '\nДобавлено нормативов: ' + res.added + '.' : '') +
    '\nЦвета обновляются сами при каждой правке результатов.' +
    '\nДлинный бег записывается как «мин:сек»: 12,40 → 12:40.' +
    '\nПороги — на листе «' + NORMS_SHEET + '», их можно править.');
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

function findRowByUin(sheet, uin) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var col = sheet.getRange(2, COL_UIN, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() === uin) return i + 2;
  }
  return 0;
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
