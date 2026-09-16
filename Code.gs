/**
 * Бэкенд синхронизации личных финансов.
 *
 * Установка:
 *   1. Создай новую Google-таблицу → Расширения → Apps Script.
 *   2. Вставь этот файл вместо содержимого Code.gs, сохрани.
 *   3. Запусти функцию setupSync() один раз, разреши доступ.
 *      Токен появится в журнале выполнения (Ctrl+Enter).
 *   4. Развернуть → Новое развёртывание → тип «Веб-приложение»:
 *        Выполнять от имени — «Я»
 *        У кого есть доступ — «Все»
 *      Скопируй адрес, оканчивающийся на /exec.
 *   5. Адрес и токен вставь в настройки приложения.
 *
 * Доступ ограничен токеном, а не логином Google: адрес держи при себе.
 * Лист перезаписывается целиком при каждой записи — это надёжно и просто,
 * но рассчитано на десятки тысяч строк, не на миллионы.
 */

var TZ = 'Asia/Tashkent';
var TX_SHEET = 'TX';
var DICT_SHEET = 'DICT';

var TX_COLS = ['id', 'syncedAt', 'updatedAt', 'deleted', 'kind', 'date', 'amount',
               'cur', 'rate', 'acc', 'acc2', 'amount2', 'cat', 'goal', 'note'];
var TX_TEXT = ['id', 'kind', 'date', 'cur', 'acc', 'acc2', 'cat', 'goal', 'note'];

var DICT_COLS = ['id', 'syncedAt', 'updatedAt', 'deleted', 'type', 'json'];
var DICT_TEXT = ['id', 'type', 'json'];

/* ---------------------------------------------------------
   Однократная настройка
   --------------------------------------------------------- */
function setupSync() {
  sheet_(TX_SHEET, TX_COLS, TX_TEXT);
  sheet_(DICT_SHEET, DICT_COLS, DICT_TEXT);
  var p = PropertiesService.getScriptProperties();
  var t = p.getProperty('SYNC_TOKEN');
  if (!t) {
    t = Utilities.getUuid().replace(/-/g, '');
    p.setProperty('SYNC_TOKEN', t);
  }
  Logger.log('SYNC_TOKEN: ' + t);
  return t;
}

/* ---------------------------------------------------------
   Точки входа веб-приложения
   --------------------------------------------------------- */

/** Отдаёт записи, изменённые на сервере после отметки since */
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (!auth_(p.token)) return json_({ ok: false, error: 'unauthorized' });
    var since = Number(p.since || 0) || 0;
    return json_({
      ok: true,
      now: peek_(),
      tx: rowsSince_(TX_SHEET, TX_COLS, TX_TEXT, since),
      dict: rowsSince_(DICT_SHEET, DICT_COLS, DICT_TEXT, since)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Принимает локальные изменения клиента */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!auth_(body.token)) return json_({ ok: false, error: 'unauthorized' });
    lock.waitLock(25000);
    var now = stamp_();
    var a = apply_(TX_SHEET, TX_COLS, TX_TEXT, body.tx || [], now);
    var b = apply_(DICT_SHEET, DICT_COLS, DICT_TEXT, body.dict || [], now);
    return json_({ ok: true, now: now, applied: a + b });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ---------------------------------------------------------
   Служебное
   --------------------------------------------------------- */
/**
 * Монотонный штамп записи.
 * Курсор выборки идёт по syncedAt со строгим «больше»; если два обмена
 * попадут в одну миллисекунду, устройство навсегда пропустит чужие правки.
 * Поэтому штамп всегда строго больше предыдущего, даже при совпадении часов.
 */
function stamp_() {
  var p = PropertiesService.getScriptProperties();
  var last = Number(p.getProperty('LAST_STAMP') || 0);
  var now = Math.max(Date.now(), last + 1);
  p.setProperty('LAST_STAMP', String(now));
  return now;
}

/** Последний выданный штамп: до него клиент уже всё видел */
function peek_() {
  return Number(PropertiesService.getScriptProperties().getProperty('LAST_STAMP') || 0);
}

function auth_(t) {
  var want = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  return !!want && String(t) === want;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Лист с заголовком; текстовые колонки помечаются форматом @,
    иначе Sheets превратит даты в Date, а id из цифр — в число */
function sheet_(name, cols, textCols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    for (var i = 0; i < textCols.length; i++) {
      sh.getRange(1, cols.indexOf(textCols[i]) + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    }
  }
  return sh;
}

function norm_(col, v, textCols) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (textCols.indexOf(col) >= 0) return (v === null || v === undefined) ? '' : String(v);
  return v === '' ? '' : Number(v);
}

/** Пакетное чтение всего листа одним getValues */
function readAll_(name, cols, textCols) {
  var sh = sheet_(name, cols, textCols);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var o = {};
    for (var c = 0; c < cols.length; c++) o[cols[c]] = norm_(cols[c], vals[r][c], textCols);
    if (o.id !== '') out.push(o);
  }
  return out;
}

function rowsSince_(name, cols, textCols, since) {
  var all = readAll_(name, cols, textCols);
  var out = [];
  for (var i = 0; i < all.length; i++) {
    if (Number(all[i].syncedAt || 0) > since) out.push(all[i]);
  }
  return out;
}

/**
 * Слияние входящих записей.
 * Побеждает запись с более поздним updatedAt (время клиента).
 * syncedAt проставляется временем сервера — по нему работает курсор выборки,
 * поэтому расхождение часов на устройствах не ломает пагинацию.
 * Удаление — это запись с deleted=1, строка не удаляется физически.
 */
function apply_(name, cols, textCols, incoming, now) {
  if (!incoming || !incoming.length) return 0;
  var sh = sheet_(name, cols, textCols);
  var rows = readAll_(name, cols, textCols);
  var idx = {};
  for (var i = 0; i < rows.length; i++) idx[rows[i].id] = i;

  var n = 0;
  for (var k = 0; k < incoming.length; k++) {
    var r = incoming[k];
    var id = String(r.id || '');
    if (!id) continue;
    var cur = (idx[id] != null) ? rows[idx[id]] : null;
    if (cur && Number(cur.updatedAt || 0) >= Number(r.updatedAt || 0)) continue; // серверная свежее

    var o = {};
    for (var c = 0; c < cols.length; c++) {
      var col = cols[c];
      o[col] = (r[col] === undefined || r[col] === null) ? '' : r[col];
    }
    o.id = id;
    o.syncedAt = now;
    o.updatedAt = Number(r.updatedAt || now);
    o.deleted = r.deleted ? 1 : 0;

    if (cur) rows[idx[id]] = o;
    else { idx[id] = rows.length; rows.push(o); }
    n++;
  }
  if (!n) return 0;

  var out = [];
  for (var j = 0; j < rows.length; j++) {
    var line = [];
    for (var m = 0; m < cols.length; m++) {
      var v = rows[j][cols[m]];
      line.push(v === undefined ? '' : v);
    }
    out.push(line);
  }
  sh.getRange(2, 1, out.length, cols.length).setValues(out); // одна пакетная запись
  return n;
}
