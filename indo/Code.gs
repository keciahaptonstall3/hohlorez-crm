/**
 * INDO CRM — Telegram-бот и сервер мини-приложения поверх таблицы «Заявки».
 *
 * Лист «Заявки» только читается (его может перезаписывать автокопия).
 * Вся работа менеджеров пишется в отдельный лист «CRM», доступ — в лист «Доступ».
 * Одна заявка = один номер телефона: повторные заявки с того же номера склеиваются.
 */
const CFG = {
  BOT_TOKEN: 'ВСТАВЬТЕ_ТОКЕН_ОТ_BOTFATHER',
  WEBAPP_URL: 'ВСТАВЬТЕ_ССЫЛКУ_НА_ВЕБ_ПРИЛОЖЕНИЕ_/exec',
  MINIAPP_URL: 'https://keciahaptonstall3.github.io/hohlorez-crm/indo/',
  SHEET_ID: '16gr2ExRxO2i_Tu1zvjigE2ELx8JjpVVx3-W8bbHGTrI',
  LEADS: 'Заявки', CRM: 'CRM', ACCESS: 'Доступ',
  LOCK_MIN: 15,                                   // сколько минут выданная заявка закреплена за человеком
};

const ST = { NEW: 'Новая', GIVEN: 'Выдана', WA: 'В WhatsApp', RECALL: 'Перезвон', NO: 'Отказ', BAD: 'Не подходит' };
const RES = { w: ST.WA, r: ST.RECALL, o: ST.NO, n: ST.BAD, s: ST.NEW };
const RECALL_MS = [3 * 3600e3, 24 * 3600e3, 48 * 3600e3];
const RECALL_TXT = ['Через 3 часа', 'Через сутки', 'Через 2 дня'];
const CRM_HEAD = ['Ключ', 'Имя', 'Телефон', 'Город', 'Паспорт', 'Статус', 'Кто', 'Обновлено', 'Перезвон', 'Звонков', 'Комментарии', 'Взял ID', 'Взял до'];
const C_ = { k: 0, n: 1, ph: 2, c: 3, pass: 4, st: 5, who: 6, upd: 7, rc: 8, tries: 9, cm: 10, lockId: 11, lockTo: 12 };
const BTN_NEXT = '📞 Дай заявку';

// ---------- запуск один раз из редактора ----------
function setup() {
  sheet_(CFG.CRM, CRM_HEAD);
  sheet_(CFG.ACCESS, ['Telegram ID', 'Имя']);
  Logger.log(tg_('setWebhook', { url: CFG.WEBAPP_URL, allowed_updates: ['message', 'callback_query'], drop_pending_updates: true }));
  Logger.log(tg_('setChatMenuButton', { menu_button: { type: 'web_app', text: 'CRM', web_app: { url: miniUrl_() } } }));
  Logger.log(tg_('setMyCommands', { commands: [{ command: 'next', description: 'Дай заявку' }, { command: 'start', description: 'Открыть CRM' }] }));
}

// ---------- вход ----------
function doGet() { return ContentService.createTextOutput('INDO CRM работает'); }

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (x) {}
  if (b.update_id !== undefined) {
    try { tgUpdate_(b); } catch (x) { console.error(x && x.stack || x); }
    return;   // без ответа Apps Script отдаёт 200; ContentService дал бы 302, и Telegram застревает на повторах одного апдейта
  }
  let out;
  try { out = { d: api_(auth_(b.init), b.op, b.a || {}) }; } catch (x) { out = { err: String(x && x.message || x) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function api_(u, op, a) {
  if (op === 'me') return { name: u.name };
  if (op === 'next') return next_(u);
  if (op === 'res') { const R = result_(u, String(a.k), a.r, +a.d || 0); return { res: R.res, st: R.st }; }
  if (op === 'comment') { comment_(u, String(a.k), String(a.text || '')); return 1; }
  if (op === 'list') return list_();
  if (op === 'card') return card_(String(a.k));
  throw new Error('Неизвестная команда');
}

// ---------- проверка, что запрос пришёл из Telegram от своего человека ----------
function auth_(init) {
  const p = {};
  String(init || '').split('&').forEach(kv => { const i = kv.indexOf('='); if (i > 0) p[dec_(kv.slice(0, i))] = dec_(kv.slice(i + 1)); });
  const hash = p.hash; delete p.hash;
  if (!hash || !p.user) throw new Error('auth');
  const check = Object.keys(p).sort().map(k => k + '=' + p[k]).join('\n');
  const secret = Utilities.computeHmacSha256Signature(Utilities.newBlob(CFG.BOT_TOKEN).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret);
  if (hex_(sig) !== hash) throw new Error('auth');
  if (Date.now() / 1000 - Number(p.auth_date) > 3 * 86400) throw new Error('auth_old');
  const user = JSON.parse(p.user), name = access_(user.id);
  if (!name) throw new Error('noaccess');
  return { id: String(user.id), name: name };
}
function dec_(s) { return decodeURIComponent(s.replace(/\+/g, '%20')); }
function hex_(bytes) { return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }

function access_(id) {
  const v = sheet_(CFG.ACCESS, ['Telegram ID', 'Имя']).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === String(id)) return String(v[i][1] || '').trim() || 'Менеджер';
  return null;
}

// ---------- таблица ----------
let SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(CFG.SHEET_ID)); }
function sheet_(name, head) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

// Номер → ключ из цифр с кодом страны. Чинит частые ошибки формы: +62 (081)…, +62 (628)…, 0853…
function normPhone_(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (!d) return '';
  if (d[0] === '0') d = '62' + d.slice(1);
  else if (d[0] === '8' && d.length >= 9 && d.length <= 12) d = '62' + d;
  if (d.indexOf('6262') === 0 && d.length > 13) d = d.slice(2);
  if (d.indexOf('620') === 0) d = '62' + d.slice(3);
  return d.length >= 9 ? d : '';
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function toMs_(v) {
  if (v instanceof Date) return v.getTime();
  const s = String(v || '').trim(); if (!s) return 0;
  try { return Utilities.parseDate(s, tz_(), 'yyyy-MM-dd HH:mm:ss').getTime(); } catch (x) {}
  const t = Date.parse(s); return isNaN(t) ? 0 : t;
}
function fmt_(ms) { return ms ? Utilities.formatDate(new Date(ms), tz_(), 'dd.MM HH:mm') : ''; }
function dms_(v) { return v instanceof Date ? v.getTime() : 0; }

// Все заявки, склеенные по номеру: { ключ: {k, n, ph, em, c, pass, t, rows[]} }
function leads_() {
  const sh = ss_().getSheetByName(CFG.LEADS) || ss_().getSheets()[0];
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim().toUpperCase());
  const col = (name, def) => { const i = h.indexOf(name); return i >= 0 ? i : def; };
  const cN = col('ИМЯ', 0), cE = col('ПОЧТА', 1), cP = col('ТЕЛЕФОН', 2), cT = col('ВРЕМЯ ЗАЯВКИ', 3), cX = col('ПАСПОРТ', 4), cC = col('ГОРОД', 5);
  const map = {};
  for (let i = 1; i < v.length; i++) {
    const r = v[i], k = normPhone_(r[cP]); if (!k) continue;
    const t = toMs_(r[cT]), x = String(r[cX]).trim();
    const L = map[k] || (map[k] = { k: k, n: '', ph: '', em: '', c: '', pass: '', t: -1, rows: [] });
    L.rows.push({ row: i + 1, n: String(r[cN]).trim(), c: String(r[cC]).trim(), em: String(r[cE]).trim(), ph: String(r[cP]).trim(), pass: x, t: fmt_(t) });
    if (t >= L.t) { L.t = t; L.n = String(r[cN]).trim(); L.ph = String(r[cP]).trim(); L.em = String(r[cE]).trim(); L.c = String(r[cC]).trim(); }
    if (x === 'Да') L.pass = 'Да'; else if (!L.pass && x) L.pass = x;
  }
  return map;
}

function crm_() {
  const sh = sheet_(CFG.CRM, CRM_HEAD), v = sh.getDataRange().getValues(), m = {};
  for (let i = 1; i < v.length; i++) { const k = String(v[i][0]).replace(/\D/g, ''); if (k) m[k] = { i: i + 1, r: v[i] }; }
  return { sh: sh, m: m };
}

function state_(e, now) {
  const r = e ? e.r : null;
  let st = r && r[C_.st] ? String(r[C_.st]) : ST.NEW;
  const lockTo = r ? dms_(r[C_.lockTo]) : 0;
  if (st === ST.GIVEN && lockTo <= now) st = ST.NEW;           // не закрыли вовремя — снова в очереди
  return {
    st: st, who: r ? String(r[C_.who]) : '', upd: r ? dms_(r[C_.upd]) : 0, rc: r ? dms_(r[C_.rc]) : 0,
    tries: r ? Number(r[C_.tries]) || 0 : 0, cm: r ? String(r[C_.cm] || '') : '',
    lockId: st === ST.GIVEN ? String(r[C_.lockId]) : '', lockTo: st === ST.GIVEN ? lockTo : 0,
  };
}

function save_(C, L, patch) {
  const e = C.m[L.k];
  const r = e ? e.r.slice() : [L.k, '', '', '', '', ST.NEW, '', '', '', 0, '', '', ''];
  r[C_.k] = L.k; r[C_.n] = L.n; r[C_.ph] = L.ph; r[C_.c] = L.c; r[C_.pass] = L.pass;
  Object.keys(patch).forEach(f => { r[C_[f]] = patch[f]; });
  const row = e ? e.i : C.sh.getLastRow() + 1;
  C.sh.getRange(row, 1, 1, 1).setNumberFormat('@');
  C.sh.getRange(row, 1, 1, r.length).setValues([r]);
  C.m[L.k] = { i: row, r: r };
}

// ---------- очередь «Дай заявку» ----------
// Сначала перезвоны, у которых подошло время; потом новые: свежие дни раньше, внутри дня — с паспортом раньше.
function queue_(LM, C, now) {
  const due = [], fresh = [], day = t => Math.floor(t / 864e5);
  Object.keys(LM).forEach(k => {
    const s = state_(C.m[k], now);
    if (s.st === ST.NEW) fresh.push(LM[k]);
    else if (s.st === ST.RECALL && s.rc <= now) due.push([LM[k], s.rc]);
  });
  due.sort((a, b) => a[1] - b[1]);
  fresh.sort((a, b) => day(b.t) - day(a.t) || (b.pass === 'Да') - (a.pass === 'Да') || b.t - a.t);
  return due.map(x => x[0]).concat(fresh);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function next_(u) {
  return withLock_(() => {
    const LM = leads_(), C = crm_(), now = Date.now();
    const Q = queue_(LM, C, now);
    // уже держит заявку — отдаём её же, а не новую
    for (const k in C.m) {
      const s = state_(C.m[k], now);
      if (s.st === ST.GIVEN && s.lockId === u.id && LM[k]) return { lead: out_(LM[k], s), left: Q.length };
    }
    if (!Q.length) return { lead: null, left: 0 };
    const L = Q[0], wasRecall = state_(C.m[L.k], now).st === ST.RECALL;
    save_(C, L, { st: ST.GIVEN, who: u.name, upd: new Date(now), lockId: u.id, lockTo: new Date(now + CFG.LOCK_MIN * 60e3) });
    const out = out_(L, state_(C.m[L.k], now)); out.recall = wasRecall;
    return { lead: out, left: Q.length - 1 };
  });
}

function result_(u, k, r, d) {
  if (!RES[r]) throw new Error('Неизвестный итог');
  return withLock_(() => {
    const LM = leads_(), C = crm_(), now = Date.now(), L = LM[k];
    if (!L) throw new Error('Заявка не найдена');
    const s = state_(C.m[k], now);
    const patch = { st: RES[r], who: u.name, upd: new Date(now), lockId: '', lockTo: '', rc: '' };
    if (r !== 's') patch.tries = s.tries + 1;
    let res = RES[r];
    if (r === 'r') { const at = now + (RECALL_MS[d] || RECALL_MS[0]); patch.rc = new Date(at); res += ' · ' + fmt_(at); }
    if (r === 's') res = 'Вернул в очередь';
    save_(C, L, patch);
    return { L: L, s: state_(C.m[k], now), res: res, st: RES[r] };
  });
}

function comment_(u, k, text) {
  text = text.trim(); if (text.length < 2) throw new Error('Пустой комментарий');
  return withLock_(() => {
    const LM = leads_(), C = crm_(), L = LM[k];
    if (!L) throw new Error('Заявка не найдена');
    const s = state_(C.m[k], Date.now()), line = fmt_(Date.now()) + ' ' + u.name + ': ' + text.replace(/\n+/g, ' ');
    save_(C, L, { cm: s.cm ? line + '\n' + s.cm : line });
  });
}

function out_(L, s) {
  return {
    k: L.k, n: L.n, ph: L.ph, em: L.em, c: L.c, pass: L.pass, t: fmt_(L.t), apps: L.rows.length,
    st: s.st, who: s.who, upd: fmt_(s.upd), rc: s.st === ST.RECALL ? fmt_(s.rc) : '', tries: s.tries,
    cm: s.cm ? s.cm.split('\n')[0] : '', until: s.lockTo || 0,
  };
}

function list_() {
  const LM = leads_(), C = crm_(), now = Date.now();
  return Object.keys(LM).map(k => LM[k]).sort((a, b) => b.t - a.t).map(L => out_(L, state_(C.m[L.k], now)));
}

function card_(k) {
  const LM = leads_(), C = crm_(), L = LM[k];
  if (!L) throw new Error('Заявка не найдена');
  const s = state_(C.m[k], Date.now()), o = out_(L, s);
  o.cms = s.cm ? s.cm.split('\n') : [];
  o.rows = L.rows.slice().reverse();
  return o;
}

// ---------- Telegram-бот ----------
function tg_(method, payload) {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + CFG.BOT_TOKEN + '/' + method,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  return JSON.parse(r.getContentText());
}
function miniUrl_() { return CFG.MINIAPP_URL + '?api=' + encodeURIComponent(CFG.WEBAPP_URL); }
function send_(chat, text, markup) { return tg_('sendMessage', { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true, reply_markup: markup }); }
function h_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
const MAIN_KB = { keyboard: [[{ text: BTN_NEXT }]], resize_keyboard: true, is_persistent: true };

function tgUpdate_(up) {
  const cache = CacheService.getScriptCache();
  if (cache.get('u' + up.update_id)) return;                  // Telegram иногда присылает одно и то же дважды
  cache.put('u' + up.update_id, '1', 21600);
  if (up.message) return onMsg_(up.message);
  if (up.callback_query) return onCb_(up.callback_query);
}

function who_(from) { const n = access_(from.id); return n ? { id: String(from.id), name: n } : null; }

function onMsg_(m) {
  if (!m.chat || m.chat.type !== 'private') return;
  const chat = m.chat.id, t = String(m.text || '').trim(), u = who_(m.from);
  if (!u) return send_(chat, 'Нет доступа к CRM.\nВаш Telegram ID: <code>' + m.from.id + '</code>\nДобавьте его в лист «Доступ» таблицы и нажмите /start.');
  if (/^\/start/.test(t)) {
    send_(chat, 'Привет, ' + h_(u.name) + '!\nКнопка «' + BTN_NEXT + '» внизу выдаёт следующую заявку со ссылкой на WhatsApp.', MAIN_KB);
    return send_(chat, 'Все заявки, поиск и карточки — в мини-приложении:', { inline_keyboard: [[{ text: '📱 Открыть CRM', web_app: { url: miniUrl_() } }]] });
  }
  if (t === BTN_NEXT || /^\/next/.test(t)) return giveLead_(chat, u);
  const rep = m.reply_to_message, k = rep && rep.from && rep.from.is_bot ? keyFromText_(rep.text) : '';
  if (k && t) {
    try { comment_(u, k, t); return send_(chat, '💬 Комментарий сохранён'); } catch (x) { return send_(chat, h_(x.message)); }
  }
  send_(chat, 'Нажмите «' + BTN_NEXT + '». Чтобы оставить комментарий к заявке, ответьте (reply) на сообщение с ней.', MAIN_KB);
}

function giveLead_(chat, u) {
  const q = next_(u);
  if (!q.lead) return send_(chat, '🎉 Очередь пуста: все заявки разобраны или ждут перезвона.', MAIN_KB);
  send_(chat, leadText_(q.lead) + '\n\n<i>Заявка ваша ' + CFG.LOCK_MIN + ' минут. Ответьте на это сообщение, чтобы оставить комментарий. В очереди ещё ' + q.left + '.</i>', leadKb_(q.lead.k));
}

function leadText_(L) {
  return (L.recall ? '🔁 <b>Перезвон</b>\n' : '') + '📞 <b>' + h_(L.n || 'Без имени') + '</b>\n'
    + '📍 ' + h_(L.c || '—') + ' · паспорт: ' + h_(L.pass || '—') + '\n'
    + '📱 ' + h_(L.ph) + '\n'
    + (L.em ? '✉️ ' + h_(L.em) + '\n' : '')
    + '🕒 заявка ' + h_(L.t) + (L.apps > 1 ? ' (заявок с номера: ' + L.apps + ')' : '') + ' · звонков ' + L.tries
    + (L.cm ? '\n💬 ' + h_(L.cm) : '');
}
function keyFromText_(text) {
  const m = String(text || '').match(/📱\s*([^\n]+)/);
  return m ? normPhone_(m[1]) : '';
}
function waBtn_(k) { return [{ text: '💬 Написать в WhatsApp', url: 'https://wa.me/' + k }]; }
function leadKb_(k) {
  return { inline_keyboard: [waBtn_(k),
    [{ text: '✅ В WhatsApp', callback_data: 'r:w:' + k }, { text: '🔁 Перезвон', callback_data: 'r:r:' + k }],
    [{ text: '❌ Отказ', callback_data: 'r:o:' + k }, { text: '🚫 Не подходит', callback_data: 'r:n:' + k }],
    [{ text: '↩️ Вернуть в очередь', callback_data: 'r:s:' + k }]] };
}

function onCb_(q) {
  const u = who_(q.from), msg = q.message;
  if (!u) return tg_('answerCallbackQuery', { callback_query_id: q.id, text: 'Нет доступа', show_alert: true });
  const p = String(q.data || '').split(':'), a = p[0], x = p[1], k = p[2];
  const kb = markup => tg_('editMessageReplyMarkup', { chat_id: msg.chat.id, message_id: msg.message_id, reply_markup: markup });
  try {
    if (a === 'next') { tg_('answerCallbackQuery', { callback_query_id: q.id }); if (k) kb({ inline_keyboard: [waBtn_(k)] }); return giveLead_(msg.chat.id, u); }
    if (a === 'r' && x === 'r') {
      tg_('answerCallbackQuery', { callback_query_id: q.id, text: 'Когда перезвонить?' });
      return kb({ inline_keyboard: [waBtn_(k), RECALL_TXT.map((t, i) => ({ text: t, callback_data: 'rc:' + i + ':' + k })), [{ text: '← Назад', callback_data: 'b::' + k }]] });
    }
    if (a === 'b') { tg_('answerCallbackQuery', { callback_query_id: q.id }); return kb(leadKb_(k)); }
    if (a === 'r' || a === 'rc') {
      const R = result_(u, k, a === 'rc' ? 'r' : x, a === 'rc' ? +x : 0);
      tg_('answerCallbackQuery', { callback_query_id: q.id, text: R.res });
      return tg_('editMessageText', { chat_id: msg.chat.id, message_id: msg.message_id, parse_mode: 'HTML', disable_web_page_preview: true,
        text: leadText_(out_(R.L, R.s)) + '\n\n<b>Итог: ' + h_(R.res) + '</b> (' + h_(u.name) + ')',
        reply_markup: { inline_keyboard: [waBtn_(k), [{ text: '📞 Следующая заявка', callback_data: 'next::' + k }]] } });
    }
    tg_('answerCallbackQuery', { callback_query_id: q.id });
  } catch (err) {
    tg_('answerCallbackQuery', { callback_query_id: q.id, text: String(err.message || err), show_alert: true });
  }
}
