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
// Google Таблица «База сдавших»: кнопка «В базу» дописывает туда всех участников текущей даты
var BASE_SPREADSHEET_ID = '1bXuLV58UVociZkxOQHipm1xCkaPVrBiMhicIB9QgqEg';
var BASE_SHEET = 'База сдавших';
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
  try { fillMissingNumbers(sheet); } catch (err) {}
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
    lock.waitLock(params.action === 'loadDate' || params.action === 'import' || params.action === 'toBase' ? 30000 : 10000);
    if (params.action === 'add') return jsonOut(addParticipant(params));
    if (params.action === 'loadDate') return jsonOut(loadDate(params.sheet, params.mode, params.force));
    if (params.action === 'toBase') return jsonOut(sendToBase(params.again));
    if (params.action === 'renumber') return jsonOut(sortAndRenumber());
    if (params.action === 'import') return jsonOut(importResults(params));
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
  try { fillMissingNumbers(sheet); } catch (err) {}
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
function loadDate(name, mode, force) {
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
  // Вместо архива — «База сдавших»: не даём стереть протокол, который туда ещё не отправили
  if (hasParticipants(sheet) && !sentToBase() && !force) {
    return { ok: false, notSent: true, event: currentEvent(),
             error: 'протокол «' + (currentEvent() || 'текущей даты') + '» ещё не отправлен в «Базу сдавших» — сначала нажмите «В базу»' };
  }
  var archive = null;
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

// ===================== «База сдавших» =====================
// «В базу» дописывает ВСЕХ участников «Списка» ниже уже записанных строк (ничего не перезаписывает).
// Дата сдачи — из названия даты в регистрации («23.11.26 (3-4 классы)» → 23.11.2026).
function baseState() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty('base_sent') || '{}'); } catch (e) { return {}; }
}
function sentToBase() { var st = baseState(); return !!st.event && st.event === currentEvent(); }

function eventDate(name) {
  var m = String(name || '').match(/(\d{1,2})[.\/](\d{1,2})[.\/](\d{2,4})/);
  if (!m) return null;
  var y = +m[3]; if (y < 100) y += 2000;
  return new Date(y, +m[2] - 1, +m[1]);
}

function sendToBase(again) {
  var event = currentEvent();
  var date = eventDate(event);
  if (!date) return { ok: false, error: 'в названии загруженной даты («' + (event || '—') + '») не найдена дата вида 23.11.26' };
  var st = baseState();
  if (st.event === event && !again) return { ok: false, already: true, at: st.at, count: st.count, event: event };

  var sheet = getSheet();
  var last = sheet.getLastRow();
  if (last < 2) return { ok: false, error: 'в «Списке» нет участников' };
  var shown = sheet.getRange(1, 1, last, COL_LETTER).getDisplayValues();
  var head = shown[0].map(normKeyPart);
  var people = shown.slice(1).filter(function (r) { return String(r[COL_FIO - 1]).trim() !== ''; });
  if (!people.length) return { ok: false, error: 'в «Списке» нет участников' };

  var base = SpreadsheetApp.openById(BASE_SPREADSHEET_ID);
  var bs = baseSheetFor(base, date);
  var lay = baseLayout(bs, head);
  var bw = lay.width, bhead = lay.head;
  if (lay.fio === -1) return { ok: false, error: 'на листе «' + bs.getName() + '» в первой строке нет столбца «ФИО»' };
  var asNumber = function (v) {
    var t = String(v).trim();
    return /^-?\d+([.,]\d+)?$/.test(t) ? Number(t.replace(',', '.')) : t;
  };
  var rows = people.map(function (r) {
    var cls = (String(r[COL_CLASS - 1]).trim() + ' ' + String(r[COL_LETTER - 1]).trim().toUpperCase()).trim();
    return bhead.map(function (h) {
      if (h === 'дата сдачи' || h === 'дата') return date;
      if (h === 'фио') return String(r[COL_FIO - 1]).replace(/\s+/g, ' ').trim();
      if (h === 'уин') return String(r[COL_UIN - 1]).trim();
      if (h === 'ступень') return String(r[COL_STUPEN - 1]).trim();
      if (h === 'пол') return String(r[COL_POL - 1]).trim();
      if (h === 'класс') return cls;
      var j = head.indexOf(h);
      if (j >= COL_RES_FIRST - 1 && j <= COL_RES_LAST - 1) return asNumber(r[j]);
      return '';
    });
  });

  var start = baseLastRow(bs, lay) + 1;                     // строго ниже последней записи (по ФИО — справа аналитика)
  var need = start + rows.length - 1;
  if (need > bs.getMaxRows()) bs.insertRowsAfter(bs.getMaxRows(), need - bs.getMaxRows());
  var uinCol = bhead.indexOf('уин');
  var dateCol = bhead.indexOf('дата сдачи') !== -1 ? bhead.indexOf('дата сдачи') : bhead.indexOf('дата');
  if (uinCol !== -1) bs.getRange(start, uinCol + 1, rows.length, 1).setNumberFormat('@');
  if (dateCol !== -1) bs.getRange(start, dateCol + 1, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  bs.getRange(start, 1, rows.length, bw).setValues(rows);
  SpreadsheetApp.flush();
  try { refreshBaseSheet(bs); } catch (e) {}                 // подсветка знаков и аналитика справа

  var at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  PropertiesService.getDocumentProperties().setProperty('base_sent', JSON.stringify({ event: event, at: at, count: rows.length }));
  var ctl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET);
  if (ctl) ctl.getRange(4, 1, 1, 2).setValues([['В базу отправлено', at + ' — ' + rows.length + ' чел.']]);
  return { ok: true, count: rows.length, event: event, sheet: bs.getName(),
           date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy'), at: at };
}

// Лист базы по учебному году даты сдачи: «База сдавших 2026-2027» (сентябрь–август).
// Нет листа этого года — пишем в лист самого позднего года; нет листов с годами — в первый.
function baseSheetFor(base, date) {
  var y = date.getFullYear(), start = date.getMonth() >= 8 ? y : y - 1;
  var best = null, bestYear = -1;
  var sheets = base.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(/^\s*база сдавших\s*(\d{4})\s*[-–—]\s*(\d{4})/i);
    if (!m) continue;
    if (+m[1] === start) return sheets[i];
    if (+m[1] > bestYear) { bestYear = +m[1]; best = sheets[i]; }
  }
  return best || base.getSheetByName(BASE_SHEET) || sheets[0];
}

function isBaseSheet(sh) { return /^\s*база сдавших/i.test(sh.getName()); }

// Столбцы данных базы — подряд с A, пока заголовок знакомый (дальше — аналитика)
function baseLayout(bs, listHead) {
  var cols = Math.max(bs.getLastColumn(), 1);
  var raw = bs.getRange(1, 1, 1, cols).getDisplayValues()[0];
  var known = { 'дата сдачи': 1, 'дата': 1, 'фио': 1, 'уин': 1, 'ступень': 1, 'пол': 1, 'класс': 1 };
  var disc = {};
  (listHead || getSheet().getRange(1, COL_RES_FIRST, 1, COL_RES_LAST - COL_RES_FIRST + 1).getDisplayValues()[0].map(normKeyPart))
    .forEach(function (h) { if (h) disc[h] = 1; });
  var head = [];
  for (var c = 0; c < raw.length; c++) {
    var h = normKeyPart(raw[c]);
    if (!h || !(known[h] || disc[h])) break;
    head.push(h);
  }
  var at = function (n) { return head.indexOf(n); };
  return { width: Math.max(head.length, 1), head: head, fio: at('фио'), st: at('ступень'), pol: at('пол'), cls: at('класс'), uin: at('уин'),
           date: at('дата сдачи') !== -1 ? at('дата сдачи') : at('дата'),
           res: head.map(function (h, i) { return disc[h] ? i : -1; }).filter(function (i) { return i !== -1; }) };
}

function baseLastRow(bs, lay) {
  var last = bs.getLastRow();
  if (last < 2 || lay.fio === -1) return 1;
  var col = bs.getRange(2, lay.fio + 1, last - 1, 1).getDisplayValues();
  for (var r = col.length; r > 0; r--) if (String(col[r - 1][0]).trim() !== '') return r + 1;
  return 1;
}

// Подсветка знаков на листе базы и аналитика справа от данных (как в выгрузке). Пересчитывается целиком.
function refreshBaseSheet(bs) {
  var lay = baseLayout(bs);
  var last = baseLastRow(bs, lay);
  var map = {};
  getNorms().forEach(function (n) { map[n.key] = n; });
  var n = last - 1;
  var data = n > 0 ? bs.getRange(2, 1, n, lay.width).getDisplayValues() : [];
  var colors = { 3: BADGE_COLORS[3], 2: BADGE_COLORS[2], 1: BADGE_COLORS[1] };

  // уровни знаков
  var levels = data.map(function (r) {
    var st = (String(r[lay.st] || '').trim().split(' ')[0] || '').toUpperCase();
    var prefix = st + '|' + normKeyPart(r[lay.pol]) + '|';
    return lay.res.map(function (c) {
      var nr = map[prefix + lay.head[c]], v = String(r[c]).trim();
      if (!nr || v === '') return { lvl: 0, has: v !== '', name: nr ? (nr.name || lay.head[c]) : '' };
      var x = normValue(v, nr.fmt), lvl = 0;
      if (x !== '') {
        var up = nr.dir === 'больше';
        var ok = function (t) { return typeof t === 'number' && (up ? x >= t : x <= t); };
        lvl = ok(nr.g) ? 3 : ok(nr.s) ? 2 : ok(nr.b) ? 1 : 0;
      }
      return { lvl: lvl, has: true, name: nr.name || lay.head[c] };
    });
  });
  if (n > 0 && lay.res.length) {
    var r0 = lay.res[0], r1 = lay.res[lay.res.length - 1];
    var rg = bs.getRange(2, r0 + 1, n, r1 - r0 + 1);
    var bgs = rg.getBackgrounds();
    var ours = {}; Object.keys(colors).forEach(function (k) { ours[colors[k].toLowerCase()] = true; });
    levels.forEach(function (row, i) {
      lay.res.forEach(function (c, k) {
        var want = row[k].lvl ? colors[row[k].lvl] : (ours[String(bgs[i][c - r0]).toLowerCase()] ? '#ffffff' : bgs[i][c - r0]);
        bgs[i][c - r0] = want;
      });
    });
    rg.setBackgrounds(bgs);
  }

  // аналитика: по нормативам, по классам, по датам
  var A0 = lay.width + 1, W = 7;
  var area = bs.getRange(1, A0, Math.max(bs.getMaxRows(), 1), W);
  area.clear();
  var out = [], fills = [];
  var put = function (row, fill) { while (row.length < W) row.push(''); out.push(row); fills.push(fill || null); };
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  var key = function (r) { return String(r[lay.uin] || '').replace(/\D/g, '') || String(r[lay.fio]).trim().toLowerCase(); };
  var dates = {};
  data.forEach(function (r) { dates[String(r[lay.date] || '').trim()] = 1; });
  put(['Аналитика — ' + bs.getName()], 'title');
  put(['обновлено ' + stamp + ' · записей ' + n + ' · дат сдачи ' + Object.keys(dates).filter(Boolean).length]);
  put([]);
  put(['По нормативам', 'Сдавали', 'Золото', 'Серебро', 'Бронза', 'Сдали на знак', 'Без знака'], 'head');
  var tests = {}, order = [];
  levels.forEach(function (row) {
    row.forEach(function (x, k) {
      if (!x.has || !x.name) return;
      var t = tests[x.name] || (tests[x.name] = { c: lay.res[k], n: 0, l: [0, 0, 0, 0] });
      if (order.indexOf(x.name) === -1) order.push(x.name);
      t.n++; t.l[x.lvl]++;
    });
  });
  order.sort(function (a, b) { return (tests[a].c - tests[b].c) || (a < b ? -1 : 1); });
  var tot = [0, 0, 0, 0, 0];
  order.forEach(function (name) {
    var t = tests[name], badge = t.l[1] + t.l[2] + t.l[3];
    put([name, t.n, t.l[3], t.l[2], t.l[1], badge, t.n - badge]);
    tot[0] += t.n; tot[1] += t.l[3]; tot[2] += t.l[2]; tot[3] += t.l[1]; tot[4] += badge;
  });
  put(['Итого', tot[0], tot[1], tot[2], tot[3], tot[4], tot[0] - tot[4]], 'total');
  put([]);
  put(['По классам', 'Человек', 'Сдавали', 'Сдали на знак', 'Золото', 'Серебро', 'Бронза'], 'head');
  var cls = {};
  data.forEach(function (r, i) {
    var c = lay.cls === -1 ? '' : String(r[lay.cls]).trim();
    if (!c) c = '(без класса)';
    var o = cls[c] || (cls[c] = { people: {}, took: {}, badge: {}, l: [0, 0, 0, 0] });
    var k = key(r), any = false, anyBadge = false;
    o.people[k] = 1;
    levels[i].forEach(function (x) { if (x.has) { any = true; o.l[x.lvl]++; if (x.lvl) anyBadge = true; } });
    if (any) o.took[k] = 1;
    if (anyBadge) o.badge[k] = 1;
  });
  Object.keys(cls).sort(function (a, b) { return (parseInt(a, 10) || 99) - (parseInt(b, 10) || 99) || (a < b ? -1 : 1); }).forEach(function (c) {
    var o = cls[c];
    put([c, Object.keys(o.people).length, Object.keys(o.took).length, Object.keys(o.badge).length, o.l[3], o.l[2], o.l[1]]);
  });
  put([]);
  put(['По датам сдачи', 'Записей', 'Результатов', 'Знаков', 'Золото', 'Серебро', 'Бронза'], 'head');
  var byDate = {}, dOrder = [];
  data.forEach(function (r, i) {
    var d = String(r[lay.date] || '').trim() || '(без даты)';
    var o = byDate[d] || (byDate[d] = { n: 0, res: 0, l: [0, 0, 0, 0] });
    if (dOrder.indexOf(d) === -1) dOrder.push(d);
    o.n++;
    levels[i].forEach(function (x) { if (x.has) { o.res++; o.l[x.lvl]++; } });
  });
  dOrder.forEach(function (d) { var o = byDate[d]; put([d, o.n, o.res, o.l[1] + o.l[2] + o.l[3], o.l[3], o.l[2], o.l[1]]); });

  if (out.length > bs.getMaxRows()) bs.insertRowsAfter(bs.getMaxRows(), out.length - bs.getMaxRows());
  var block = bs.getRange(1, A0, out.length, W);
  block.setNumberFormat('@').setValues(out.map(function (r) { return r.map(function (x) { return String(x); }); }));
  block.setHorizontalAlignment('center');
  bs.getRange(1, A0, out.length, 1).setHorizontalAlignment('left');
  fills.forEach(function (f, i) {
    if (!f) return;
    var rg = bs.getRange(i + 1, A0, 1, W);
    if (f === 'title') rg.setFontWeight('bold').setFontSize(12);
    if (f === 'total') rg.setFontWeight('bold');
    if (f === 'head') {
      rg.setFontWeight('bold').setBackground('#EEF1F6');
      var t = out[i][0];
      var gi = t === 'По нормативам' ? 2 : t === 'По классам' ? 4 : 4;
      bs.getRange(i + 1, A0 + gi, 1, 3).setBackgrounds([[colors[3], colors[2], colors[1]]]);
    }
  });
  bs.setColumnWidth(A0, 300);
  for (var w = 1; w < W; w++) bs.setColumnWidth(A0 + w, 95);
  return { rows: n };
}

function menuRefreshBase() {
  var base = SpreadsheetApp.openById(BASE_SPREADSHEET_ID);
  var done = [];
  base.getSheets().forEach(function (sh) { if (isBaseSheet(sh)) { refreshBaseSheet(sh); done.push(sh.getName()); } });
  SpreadsheetApp.getUi().alert(done.length ? 'Обновлены подсветка и аналитика: ' + done.join(', ') : 'В «Базе сдавших» нет листов, начинающихся с «База сдавших».');
}

function menuToBase() {
  var ui = SpreadsheetApp.getUi();
  var event = currentEvent();
  var st = baseState();
  var again = false;
  if (st.event === event) {
    if (ui.alert('В базу', 'Протокол «' + event + '» уже отправлен в «Базу сдавших» ' + st.at + ' (' + st.count + ' чел.).\n' +
      'Отправить ещё раз? В базе появятся повторные строки.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    again = true;
  } else if (ui.alert('В базу', 'Отправить всех участников протокола «' + event + '» в «Базу сдавших»?\n' +
      'Строки допишутся ниже уже записанных, ничего не перезаписывается.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var res;
  try { res = sendToBase(again); } finally { lock.releaseLock(); }
  ui.alert(res.ok ? 'Готово: в «' + res.sheet + '» записано ' + res.count + ' чел. с датой ' + res.date + '.' : 'Не отправлено: ' + res.error);
}

// ===================== Загрузка выгрузки с другой площадки =====================
// Сайт присылает участников из чужого файла выгрузки: { fio, uin, stupen, pol, klass, letter, present,
// res: { 'отжимание': '43', 'длинная': '9:45', … } }. Участника ищем по УИН, затем по ФИО;
// кого нет — дописываем в «Список». Результат пишем в столбец B листа норматива (как судья),
// прежний результат заменяется. dry: true — только посчитать, ничего не меняя.
function importResults(params) {
  var items = params.items || [];
  var dry = !!params.dry;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheet();
  fillMissingNumbers(sheet);
  var fioKey = function (f) { return String(f || '').replace(/\s+/g, ' ').trim().toLowerCase().replace(/ё/g, 'е'); };
  var uinKey = function (u) { return String(u || '').replace(/\D/g, ''); };

  // Участники «Списка»
  var last = sheet.getLastRow(), byUin = {}, byFio = {}, maxNum = 0, lastFioRow = 1;
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, COL_UIN).getValues().forEach(function (r, i) {
      var n = parseInt(r[COL_NUM - 1], 10);
      if (n > maxNum) maxNum = n;
      if (String(r[COL_FIO - 1]).trim() === '') return;
      lastFioRow = i + 2;
      var u = uinKey(r[COL_UIN - 1]), f = fioKey(r[COL_FIO - 1]);
      if (u) byUin[u] = i + 2;
      byFio[f] = byFio[f] ? -1 : i + 2;                     // -1 — тёзки, по ФИО не сопоставляем
    });
  }

  // Столбец «Списка» → лист норматива (из формулы вида ='прыжки'!B2)
  var width = COL_RES_LAST - COL_RES_FIRST + 1;
  var head = sheet.getRange(1, COL_RES_FIRST, 1, width).getDisplayValues()[0];
  var forms = sheet.getRange(2, COL_RES_FIRST, 1, width).getFormulas()[0];
  var target = {};
  head.forEach(function (h, j) {
    var m = String(forms[j]).match(/^=\s*'?([^'!]+)'?!/);
    var sh = m && ss.getSheetByName(m[1]);
    if (h && sh) target[normKeyPart(h)] = sh;
  });

  // Сопоставление
  var stat = { items: items.length, found: 0, added: 0, results: 0, changed: 0, same: 0, present: 0, unknown: {}, newNames: [] };
  var plan = [], fresh = [];
  items.forEach(function (it) {
    var u = uinKey(it.uin), row = (u && byUin[u]) || 0;
    if (!row) { var f = byFio[fioKey(it.fio)]; if (f > 0) row = f; }
    if (row) stat.found++;
    else {
      if (String(it.fio || '').trim() === '') return;
      fresh.push(it);
      stat.added++;
      if (stat.newNames.length < 15) stat.newNames.push(String(it.fio).trim());
    }
    plan.push({ it: it, row: row });
  });

  if (!dry && fresh.length) {
    var start = lastFioRow + 1;
    writePeople(sheet, fresh.map(function (it, k) {
      return { num: maxNum + k + 1, fio: String(it.fio).replace(/\s+/g, ' ').trim(), uin: String(it.uin || '').trim(),
               stupen: String(it.stupen || '').trim(), pol: String(it.pol || '').trim(),
               klass: String(it.klass || '').trim(), letter: String(it.letter || '').trim() };
    }), start);
    var k = 0;
    plan.forEach(function (p) { if (!p.row) p.row = start + k++; });
    SpreadsheetApp.flush();
  }

  // Результаты: по листам норматива, столбец B
  var writes = {};                                          // имя листа → { row: значение }
  var current = {};
  plan.forEach(function (p) {
    Object.keys(p.it.res || {}).forEach(function (col) {
      var v = String(p.it.res[col] == null ? '' : p.it.res[col]).trim();
      if (v === '') return;
      var sh = target[normKeyPart(col)];
      if (!sh) { stat.unknown[col] = true; return; }
      var name = sh.getName();
      if (normName(name) === LONG_SHEET) v = toMinSecText(v) || v;
      stat.results++;
      if (!p.row) { stat.changed++; return; }               // новый участник (только при проверке)
      if (!current[name]) {
        var n = Math.max(sh.getLastRow(), 1);
        current[name] = n >= 2 ? sh.getRange(2, 2, n - 1, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : [];
      }
      var old = current[name][p.row - 2] || '';
      if (old === v) { stat.same++; return; }
      stat.changed++;
      (writes[name] = writes[name] || {})[p.row] = v;
    });
    if (p.it.present) stat.present++;
  });
  stat.unknown = Object.keys(stat.unknown);
  if (dry) return { ok: true, dry: true, stat: stat };

  Object.keys(writes).forEach(function (name) {
    var sh = ss.getSheetByName(name), rows = Object.keys(writes[name]).map(Number).sort(function (a, b) { return a - b; });
    var i = 0;
    while (i < rows.length) {                               // подряд идущие строки — одним блоком
      var j = i;
      while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
      var block = rows.slice(i, j + 1).map(function (r) { return [writes[name][r]]; });
      sh.getRange(rows[i], 2, block.length, 1).setNumberFormat('@').setValues(block);
      i = j + 1;
    }
  });
  plan.forEach(function (p) { if (p.it.present && p.row) sheet.getRange(p.row, COL_PRESENT).setValue(true); });
  SpreadsheetApp.flush();
  try { paintBadges(); } catch (e) {}
  return { ok: true, stat: stat };
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
  if (count !== null) sh.getRange(3, 1, 2, 2).setValues([['Участников', count], ['В базу отправлено', 'ещё нет']]);
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
    if (name === SHEET_NAME && e.range.getColumn() <= COL_FIO) fillMissingNumbers(e.range.getSheet());
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
    .addItem('В базу: отправить протокол в «Базу сдавших»', 'menuToBase')
    .addItem('База сдавших: обновить подсветку и аналитику', 'menuRefreshBase')
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
    '<p style="color:#666;font-size:12px">«Загрузить» — «Список», явка и результаты на листах нормативов очистятся ' +
    'и загрузятся участники выбранной даты. Сначала отправьте текущий протокол в «Базу сдавших» (меню ГТО → В базу).<br>' +
    '«Догрузить новых» — добавит только тех, кого ещё нет в «Списке» (по УИН), ничего не очищая.</p>' +
    '<button onclick="go(\'replace\')" style="padding:8px 12px">Загрузить</button> ' +
    '<button onclick="go(\'append\')" style="padding:8px 12px">Догрузить новых</button>' +
    '<p id="m" style="margin-top:12px"></p></div>' +
    '<script>function go(mode,force){var d=document.getElementById("d").value,m=document.getElementById("m");' +
    'if(mode==="replace"&&!force&&!confirm("Загрузить «"+d+"»?\\nСписок, явка и результаты очистятся."))return;' +
    'm.textContent="Загружаем…";document.querySelectorAll("button").forEach(function(b){b.disabled=true});' +
    'google.script.run.withSuccessHandler(function(r){document.querySelectorAll("button").forEach(function(b){b.disabled=false});' +
    'if(!r.ok&&r.notSent){m.textContent="";if(confirm("Протокол «"+r.event+"» ещё не отправлен в «Базу сдавших».\\nЕсли загрузить новую дату, его данные пропадут.\\n\\nВсё равно загрузить?"))go(mode,true);return;}' +
    'm.textContent=r.ok?(r.mode==="append"?"Добавлено: "+r.added+", уже были: "+r.skipped:"Загружено участников: "+r.loaded):("Ошибка: "+r.error)})' +
    '.withFailureHandler(function(e){m.textContent="Ошибка: "+e.message;document.querySelectorAll("button").forEach(function(b){b.disabled=false})})' +
    '.menuLoad(d,mode,!!force)}</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(330), 'Загрузка даты из регистрации');
}

function menuLoad(name, mode, force) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return loadDate(name, mode, force); } finally { lock.releaseLock(); }
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

// Участникам, вставленным в «Список» вручную без номера, даём следующие номера (max+1, max+2…),
// иначе сайт их не показывает, а на листах нормативов у них нет номера.
function fillMissingNumbers(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var ab = sheet.getRange(2, COL_NUM, last - 1, 2).getValues();
  var need = [];
  ab.forEach(function (r, i) {
    if (String(r[1]).trim() !== '' && String(r[0]).trim() === '') need.push(i);
  });
  if (!need.length) return 0;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return 0;
  try {
    ab = sheet.getRange(2, COL_NUM, last - 1, 2).getValues();   // перечитываем под замком
    var max = 0;
    ab.forEach(function (r) { var n = parseInt(r[0], 10); if (n > max) max = n; });
    // пишем только пустые ячейки, подряд идущие — одним блоком
    var added = 0, start = -1, block = [];
    var flush = function () { if (block.length) sheet.getRange(2 + start, COL_NUM, block.length, 1).setValues(block); block = []; start = -1; };
    ab.forEach(function (r, i) {
      if (String(r[1]).trim() !== '' && String(r[0]).trim() === '') {
        if (start === -1) start = i;
        block.push([++max]); added++;
      } else flush();
    });
    flush();
    if (added) SpreadsheetApp.flush();
    return added;
  } finally {
    lock.releaseLock();
  }
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
