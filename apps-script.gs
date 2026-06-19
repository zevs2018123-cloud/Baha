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
 *
 * ДОСТУП К ДАШБОРДУ (whitelist хранится на сервере, общий для всех):
 *  • Список админов лежит в Script Properties (ключ ADMINS, JSON-массив).
 *  • Первичная выдача: Project Settings → Script Properties →
 *    добавьте свойство SEED_ADMINS со значением вашего юзернейма или ID,
 *    например «@Desanji» (можно несколько через запятую). При первом
 *    обращении этот список станет начальным набором админов.
 *  • Дальше админы выдают/забирают доступ кнопкой прямо в дашборде.
 *
 * Примечание о безопасности: личность запрашивающего (byId/byName)
 * приходит с клиента и не верифицируется подписью Telegram — это тот же
 * уровень доверия, что и во всём приложении. При необходимости строгой
 * проверки можно валидировать подпись Telegram-логина по токену бота.
 */

var SHEET_NAME = 'events';
var HEADERS = ['ts', 'iso', 'session', 'uid', 'uname', 'src', 'event', 'props'];
var PROP_ = PropertiesService.getScriptProperties();

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
  }
  return sh;
}

function ok_() { return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT); }

// Запись одной строки события (с блокировкой)
function logEvent_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    getSheet_().appendRow([
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
}

// Приём: либо событие воронки с сайта, либо апдейт Telegram (вебхук бота)
function doPost(e) {
  var data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) {}
  if (data && (data.update_id !== undefined || data.message || data.edited_message || data.callback_query)) {
    return handleTelegramUpdate_(data, e); // /start от бота — собираем ID без сайта
  }
  logEvent_(data);
  return ok_();
}

/* ── Бот: обработка /start (сбор ID без веб-логина) ──────────────────── */
function handleTelegramUpdate_(update, e) {
  try {
    var need = PROP_.getProperty('WEBHOOK_SECRET');
    if (need && (!e || !e.parameter || e.parameter.wsecret !== need)) return ok_();
    var from = (update.message && update.message.from) ||
               (update.edited_message && update.edited_message.from) ||
               (update.callback_query && update.callback_query.from);
    if (!from) return ok_();
    var uid = String(from.id);
    var uname = from.username || '';
    var name = [from.first_name, from.last_name].filter(Boolean).join(' ');
    // пишем как событие — пользователь сразу появится в списке дашборда
    logEvent_({ session: 'bot', uid: uid, uname: uname, src: 'bot', event: 'bot_start', props: { name: name } });
    // верификация пройдена → отвечаем кнопкой-ссылкой на квиз
    var token = PROP_.getProperty('BOT_TOKEN');
    if (token) {
      var webapp = PROP_.getProperty('WEBAPP_URL');
      var payload = {
        chat_id: from.id,
        text: 'Привет, ' + (from.first_name || 'трейдер') + '! 👋\n\n' +
              'Ты в Школе трейдинга Бахи. Доступ открыт — жми кнопку ' +
              '«🎓 Пройти обучение» внизу, чтобы начать.'
      };
      if (webapp) {
        // reply-кнопка «выскакивает» снизу и открывает квиз внутри Telegram
        payload.reply_markup = {
          keyboard: [[{ text: '🎓 Пройти обучение', web_app: { url: webapp } }]],
          resize_keyboard: true
        };
      }
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
    }
  } catch (err) {}
  return ok_();
}

/* ── Разовая настройка бота: вебхук + кнопка-меню (Mini App) ─────────────
 * Перед запуском задайте Script Properties:
 *   BOT_TOKEN   — токен бота от @BotFather (обязателен)
 *   WEBAPP_URL  — URL сайта, напр. https://zevs2018123-cloud.github.io/Baha/
 *   WEBHOOK_SECRET — произвольная строка (необязательно, для защиты вебхука)
 * Затем выберите функцию setupBot и нажмите Run. Запускать после Deploy.
 */
function setupBot() {
  var token = PROP_.getProperty('BOT_TOKEN');
  if (!token) throw new Error('Сначала задайте Script Property BOT_TOKEN');
  var url = ScriptApp.getService().getUrl();
  var secret = PROP_.getProperty('WEBHOOK_SECRET') || '';
  if (secret) url += (url.indexOf('?') === -1 ? '?' : '&') + 'wsecret=' + encodeURIComponent(secret);
  var out = [];
  out.push('setWebhook: ' + UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + token + '/setWebhook?url=' + encodeURIComponent(url),
    { muteHttpExceptions: true }).getContentText());
  var webapp = PROP_.getProperty('WEBAPP_URL');
  if (webapp) {
    out.push('setChatMenuButton: ' + UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + token + '/setChatMenuButton', {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ menu_button: { type: 'web_app', text: 'Открыть школу', web_app: { url: webapp } } }),
        muteHttpExceptions: true }).getContentText());
  }
  Logger.log(out.join('\n'));
  return out.join('\n');
}

// Агрегаты и управление доступом для дашборда (JSONP)
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'funnel';
  var callback = p.callback;
  var payload;
  if (action === 'ping') payload = { ok: true, admins: getAdmins_(), hasToken: !!PROP_.getProperty('BOT_TOKEN'), webapp: PROP_.getProperty('WEBAPP_URL') || '' };
  else if (action === 'funnel') payload = buildFunnel_();
  else if (action === 'users') payload = buildUsers_();
  else if (action === 'access') payload = { admins: getAdmins_() };
  else if (action === 'grant' || action === 'revoke') payload = mutateAccess_(action, p);
  else payload = { error: 'unknown action' };
  var json = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ── Управление доступом к дашборду ──────────────────────────────────── */
function norm_(x) { return String(x == null ? '' : x).trim().toLowerCase(); }

function getAdmins_() {
  var list = [];
  try { list = JSON.parse(PROP_.getProperty('ADMINS') || '[]') || []; } catch (e) {}
  if (!list.length) { // первичный набор из SEED_ADMINS
    var seed = PROP_.getProperty('SEED_ADMINS') || '';
    seed.split(/[\s,]+/).forEach(function (x) { x = x.trim(); if (x) list.push(x); });
    if (list.length) PROP_.setProperty('ADMINS', JSON.stringify(list));
  }
  return list;
}

function isAdminEntry_(list, id, uname) {
  id = norm_(id);
  var un = uname ? '@' + norm_(uname).replace(/^@/, '') : '';
  for (var i = 0; i < list.length; i++) {
    var t = norm_(list[i]);
    if (!t) continue;
    if (id && t === id) return true;
    if (un && (t === un || t === un.slice(1))) return true;
  }
  return false;
}

function mutateAccess_(action, p) {
  var admins = getAdmins_();
  if (!isAdminEntry_(admins, p.byId, p.byName)) {
    return { ok: false, error: 'forbidden', admins: admins };
  }
  var target = String(p.target || '').trim();
  if (!target) return { ok: false, error: 'no target', admins: admins };
  var tn = norm_(target).replace(/^@/, '');
  admins = admins.filter(function (x) { return norm_(x).replace(/^@/, '') !== tn; });
  if (action === 'grant') admins.push(target);
  PROP_.setProperty('ADMINS', JSON.stringify(admins));
  return { ok: true, admins: admins };
}

/* ── Список собранных пользователей ──────────────────────────────────── */
function buildUsers_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { users: [], scope: 'global', updated: Date.now() };
  var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var ts = Number(rows[i][0]) || 0;
    var uid = String(rows[i][3]);
    var uname = String(rows[i][4]);
    var src = String(rows[i][5]);
    var u = map[uid] = map[uid] || { uid: uid, uname: '', src: src, firstSeen: ts, lastSeen: ts, events: 0 };
    if (uname) u.uname = uname;
    u.events++;
    if (ts && ts < u.firstSeen) u.firstSeen = ts;
    if (ts >= u.lastSeen) { u.lastSeen = ts; if (src) u.src = src; }
  }
  var users = Object.keys(map).map(function (k) { return map[k]; });
  users.sort(function (a, b) { return b.lastSeen - a.lastSeen; });
  return { users: users, scope: 'global', updated: Date.now() };
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
