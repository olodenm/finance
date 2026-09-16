/**
 * Построитель витрин для Looker Studio.
 * Читает листы TX и DICT, которые пишет sync, и собирает три плоских листа:
 *
 *   FLAT     — 1 строка = операция, с раскрытыми названиями и признаками
 *   SPINE    — 1 строка = КАЛЕНДАРНЫЙ день, включая дни без операций
 *   METRICS  — 1 строка = месяц, предагрегаты (медиана, HHI, run-rate)
 *
 * Почему SPINE обязателен: Looker Studio не умеет создавать отсутствующие даты.
 * Без хребта метрика «дни без трат» невозможна в принципе, медиана дня завышается
 * (по реальным данным — 113 500 вместо 47 575), а линия капитала рвётся на пустых датах.
 *
 * Почему METRICS отдельно: в Looker Studio НЕТ агрегата MEDIAN — ни в scorecard,
 * ни в вычисляемом поле. Медиану и HHI считаем здесь и кладём готовыми числами.
 *
 * Установка:
 *   1. В том же проекте Apps Script создать файл Report.gs и вставить этот код.
 *   2. Один раз запустить setupReport() — построит листы и повесит триггер (раз в час).
 *   3. В Looker Studio подключить источник Google Sheets → нужный лист.
 *
 * Взнос в цель НЕ считается расходом: это перемещение капитала.
 * capital = свободные средства + деньги в целях, и от взноса он не меняется.
 */

var TZ_REPORT = 'Asia/Tashkent';
var SH_FLAT = 'FLAT';
var SH_SPINE = 'SPINE';
var SH_METRICS = 'METRICS';

/* ─────────────────────────── точка входа ─────────────────────────── */

function setupReport() {
  buildReport();
  var have = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'buildReport';
  });
  if (!have) {
    ScriptApp.newTrigger('buildReport').timeBased().everyHours(1).create();
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Витрины собраны, триггер поставлен', 'Report', 5);
}

function buildReport() {
  var src = readSource_();
  writeSheet_(SH_FLAT, buildFlat_(src));
  var spine = buildSpine_(src);
  writeSheet_(SH_SPINE, spine);
  writeSheet_(SH_METRICS, buildMetrics_(src, spine));
}

/* ──────────────────────── чтение TX и DICT ───────────────────────── */

/**
 * Один getValues на лист — построчных обращений нет.
 * Записи с deleted=1 отбрасываются: это тумбстоны синхронизации.
 */
function readSource_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tx = readTable_(ss.getSheetByName('TX'));
  var dict = readTable_(ss.getSheetByName('DICT'));

  var cats = {}, accs = {}, goals = {}, settings = {};
  dict.forEach(function (r) {
    if (String(r.deleted) === '1' || !r.json) return;
    var o;
    try { o = JSON.parse(r.json); } catch (e) { return; }
    if (r.type === 'cat') cats[r.id] = o;
    else if (r.type === 'account') accs[r.id] = o;
    else if (r.type === 'goal') goals[r.id] = o;
    else if (r.type === 'settings') settings = o;
  });

  var rows = tx.filter(function (r) {
    return String(r.deleted) !== '1' && r.date && r.amount !== '' && r.amount != null;
  }).map(function (r) {
    return {
      id: r.id, kind: r.kind, date: asDate_(r.date),
      amount: Number(r.amount) || 0, cur: r.cur || 'UZS',
      rate: Number(r.rate) || 0, acc: r.acc, acc2: r.acc2,
      amount2: r.amount2 === '' ? null : Number(r.amount2),
      cat: r.cat, goal: r.goal, note: r.note || '',
      updatedAt: Number(r.updatedAt) || 0
    };
  }).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return { tx: rows, cats: cats, accs: accs, goals: goals, settings: settings };
}

function readTable_(sh) {
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  var head = vals[0].map(String);
  return vals.slice(1).map(function (row) {
    var o = {};
    for (var i = 0; i < head.length; i++) o[head[i]] = row[i];
    return o;
  });
}

/** Даты приходят то строкой, то Date — приводим к 'yyyy-MM-dd' */
function asDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ_REPORT, 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

var BASE_CUR = 'UZS';

/**
 * Текущий курс валюты к базовой из настроек.
 * Приложение хранит карту settings.rates = {USD: .., KZT: ..}; плоское
 * settings.rate осталось только для доллара и старых записей.
 */
function rateOf_(cur, settings) {
  if (!cur || cur === BASE_CUR) return 1;
  var rates = (settings && settings.rates) || {};
  var r = Number(rates[cur]);
  if (r > 0) return r;
  if (cur === 'USD') { r = Number(settings && settings.rate); if (r > 0) return r; }
  return 1;
}

/**
 * Сумма в UZS. Курс берётся ИЗ САМОЙ ОПЕРАЦИИ, а не из настроек:
 * иначе смена курса задним числом перепишет всю историю капитала.
 * Проверка «валюта не базовая» вместо прежней «валюта не доллар» —
 * с тенге старый вариант отдавал 14 500 вместо 348 000, то есть
 * занижал строку ровно в курс раз и молча.
 */
function uzs_(t, settings) {
  if (!t.cur || t.cur === BASE_CUR) return t.amount;
  var rate = Number(t.rate) > 0 ? Number(t.rate) : rateOf_(t.cur, settings);
  return t.amount * rate;
}

/* ──────────────────────────── FLAT ───────────────────────────────── */

function buildFlat_(src) {
  var out = [[
    'tx_id', 'date', 'year_month', 'dow', 'is_weekend', 'kind', 'kind_ru',
    'category', 'cat_kind', 'is_fixed', 'budget', 'account', 'currency',
    'amount', 'amount_uzs', 'goal_name', 'note'
  ]];
  var KIND_RU = { expense: 'Расход', income: 'Доход', goal: 'В цель', transfer: 'Перевод' };

  src.tx.forEach(function (t) {
    var c = src.cats[t.cat] || {};
    var g = src.goals[t.goal] || {};
    var d = new Date(t.date + 'T00:00:00');
    var dow = d.getDay() === 0 ? 7 : d.getDay();
    out.push([
      t.id, t.date, t.date.slice(0, 7), dow, dow >= 6 ? 1 : 0,
      t.kind, KIND_RU[t.kind] || t.kind,
      t.kind === 'goal' ? ('→ ' + (g.name || '')) : (c.name || '— без категории —'),
      c.kind || '', c.fixed ? 1 : 0, Number(c.budget) || 0,
      (src.accs[t.acc] || {}).name || '', t.cur,
      t.amount, Math.round(uzs_(t, src.settings)),
      g.name || '', t.note
    ]);
  });
  return out;
}

/* ──────────────────────────── SPINE ──────────────────────────────── */

/**
 * Календарный хребет от первой операции до сегодня.
 * cum_free — свободные средства, cum_capital — вместе с деньгами в целях.
 */
function buildSpine_(src) {
  var out = [[
    'date', 'year_month', 'dow', 'is_weekend', 'expense_uzs', 'income_uzs',
    'goal_uzs', 'is_zero_day', 'tx_count', 'cum_free', 'cum_capital'
  ]];
  if (!src.tx.length) return out;

  var byDay = {};
  src.tx.forEach(function (t) {
    var d = byDay[t.date] || (byDay[t.date] = { e: 0, i: 0, g: 0, n: 0 });
    var v = uzs_(t, src.settings);
    if (t.kind === 'expense') d.e += v;
    else if (t.kind === 'income') d.i += v;
    else if (t.kind === 'goal') d.g += v;
    d.n++;
  });

  var opening = 0;
  Object.keys(src.accs).forEach(function (k) {
    var a = src.accs[k];
    opening += (Number(a.opening) || 0) * rateOf_(a.cur, src.settings);
  });

  var cur = new Date(src.tx[0].date + 'T00:00:00');
  var end = new Date();
  var free = opening, pot = 0;

  while (cur <= end) {
    var key = Utilities.formatDate(cur, TZ_REPORT, 'yyyy-MM-dd');
    var r = byDay[key] || { e: 0, i: 0, g: 0, n: 0 };
    free += r.i - r.e - r.g;
    pot += r.g;
    var dow = cur.getDay() === 0 ? 7 : cur.getDay();
    out.push([
      key, key.slice(0, 7), dow, dow >= 6 ? 1 : 0,
      Math.round(r.e), Math.round(r.i), Math.round(r.g),
      r.e === 0 ? 1 : 0, r.n, Math.round(free), Math.round(free + pot)
    ]);
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

/* ─────────────────────────── METRICS ─────────────────────────────── */

function buildMetrics_(src, spine) {
  var out = [[
    'year_month', 'days', 'expense_uzs', 'income_uzs', 'goal_uzs', 'saldo_uzs',
    'median_day', 'median_active_day', 'mean_day', 'skew', 'zero_days',
    'active_days', 'tx_count', 'avg_check', 'fixed_uzs', 'fixed_share',
    'income_sources', 'income_hhi', 'run_rate_30'
  ]];

  var months = {};
  for (var i = 1; i < spine.length; i++) {
    var r = spine[i];
    (months[r[1]] = months[r[1]] || []).push({ exp: r[4], inc: r[5], goal: r[6] });
  }

  var byMonthTx = {};
  src.tx.forEach(function (t) {
    (byMonthTx[t.date.slice(0, 7)] = byMonthTx[t.date.slice(0, 7)] || []).push(t);
  });

  Object.keys(months).sort().forEach(function (m) {
    var days = months[m];
    var all = days.map(function (d) { return d.exp; });
    var active = all.filter(function (v) { return v > 0; });
    var exp = sum_(all), inc = sum_(days.map(function (d) { return d.inc; }));
    var gl = sum_(days.map(function (d) { return d.goal; }));

    var txs = byMonthTx[m] || [];
    var expTx = txs.filter(function (t) { return t.kind === 'expense'; });
    var fixed = 0;
    expTx.forEach(function (t) {
      if ((src.cats[t.cat] || {}).fixed) fixed += uzs_(t, src.settings);
    });

    // концентрация дохода: 1.0 = всё из одного источника
    var bySrc = {};
    txs.forEach(function (t) {
      if (t.kind !== 'income') return;
      bySrc[t.cat] = (bySrc[t.cat] || 0) + uzs_(t, src.settings);
    });
    var keys = Object.keys(bySrc), hhi = 0;
    keys.forEach(function (k) { if (inc) hhi += Math.pow(bySrc[k] / inc, 2); });

    var mean = all.length ? exp / all.length : 0;
    var medAll = median_(all);

    out.push([
      m, all.length, Math.round(exp), Math.round(inc), Math.round(gl),
      Math.round(inc - exp),                       // взнос в цель не вычитаем: это капитал
      Math.round(medAll), Math.round(median_(active)), Math.round(mean),
      medAll ? Math.round(mean / medAll * 100) / 100 : '',
      all.length - active.length, active.length, expTx.length,
      expTx.length ? Math.round(exp / expTx.length) : 0,
      Math.round(fixed), exp ? Math.round(fixed / exp * 1000) / 1000 : 0,
      keys.length, Math.round(hhi * 1000) / 1000,
      all.length ? Math.round(exp / all.length * 30) : 0
    ]);
  });
  return out;
}

function sum_(a) { var s = 0; for (var i = 0; i < a.length; i++) s += a[i]; return s; }

function median_(a) {
  if (!a.length) return 0;
  var x = a.slice().sort(function (p, q) { return p - q; });
  var n = x.length;
  return n % 2 ? x[(n - 1) / 2] : (x[n / 2 - 1] + x[n / 2]) / 2;
}

/* ─────────────────────────── запись листа ────────────────────────── */

/** Одна пакетная запись setValues на лист; старая область чистится целиком. */
function writeSheet_(name, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  if (!rows.length) return;
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}
