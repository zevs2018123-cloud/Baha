/**
 * Бэкенд статистики для «Школы трейдинга Бахи».
 * Принимает события воронки (doPost) и отдаёт агрегаты для дашборда (doGet, JSONP).
 *
 * УСТАНОВКА:
 *  1. Создайте новую Google-таблицу.
 *  2. Расширения → Apps Script. Вставьте этот файл целиком, сохраните.
 *  3. Deploy → New deployment → тип «Web app».
 *       Execute as: Me · Who has access: Anyone.
 *     Скопируйте URL вида https://script.google.com/macros/s/XXXX/exec
 *  4. Вставьте этот URL в index.html → CONFIG.statsEndpoint.
 *
 * Лист "events" создаётся автоматически при первом событии.
 */

var SHEET_NAME = 'events';
var HEADERS = ['ts', 'iso', 'session', 'uid', 'uname', 'src', 'event', 'props'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
  }
  return sh;
}

// Запись события
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    var data = {};
    try { data = JSON.parse(e.postData.contents); } catch (err) {}
    var sh = getSheet_();
    sh.appendRow([
      data.ts || Date.now(),
      new Date(data.ts || Date.now()).toISOString(),
      String(data.session || ''),
      String(data.uid || 'anon'),
      String(data.uname || ''),
      String(data.src || ''),
      String(data.event || ''),
      JSON.stringify(data.props || {})
    ]);
  } catch (err) {
    // молча — клиент отправляет fire-and-forget
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

// Агрегаты для дашборда (JSONP)
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'funnel';
  var callback = e && e.parameter && e.parameter.callback;
  var payload = (action === 'funnel') ? buildFunnel_() : { error: 'unknown action' };
  var json = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function buildFunnel_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { events: {}, modules: {}, totalUsers: 0, scope: 'global', updated: Date.now() };

  var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var usersByEvent = {};   // event -> {uid:1}
  var countByEvent = {};   // event -> n
  var modules = {};        // moduleId -> {opens, completes, sumCorrect, sumTotal}
  var allUsers = {};

  for (var i = 0; i < rows.length; i++) {
    var uid = String(rows[i][3]);
    var event = String(rows[i][6]);
    var props = {};
    try { props = JSON.parse(rows[i][7]); } catch (err) {}

    allUsers[uid] = 1;
    countByEvent[event] = (countByEvent[event] || 0) + 1;
    (usersByEvent[event] = usersByEvent[event] || {})[uid] = 1;

    var mid = props && props.moduleId;
    if (mid) {
      var m = modules[mid] = modules[mid] || { opens: 0, completes: 0, sumCorrect: 0, sumTotal: 0 };
      if (event === 'module_open') m.opens++;
      if (event === 'module_complete') {
        m.completes++;
        m.sumCorrect += Number(props.correct || 0);
        m.sumTotal += Number(props.total || 0);
      }
    }
  }

  var events = {};
  Object.keys(usersByEvent).forEach(function (k) {
    events[k] = { users: Object.keys(usersByEvent[k]).length, count: countByEvent[k] || 0 };
  });

  return {
    events: events,
    modules: modules,
    totalUsers: Object.keys(allUsers).length,
    scope: 'global',
    updated: Date.now()
  };
}
