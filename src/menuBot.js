/**
 * menuBot.js — интерактивное меню с авторизацией из Google Sheets.
 *
 * Листы:
 *   Logins        — Имя | Фамилия | Логин | Пароль | Tag
 *   Распределение — Объект | Управляющий
 */

const https = require('https');
const { getAuthClient, getSheetsClient } = require('./auth');
const { getConfig } = require('./config');
const { parseDateStr, todayDateKey } = require('./dateUtils');

// ── Кэш данных из таблиц ─────────────────────────────────────────────────────

let _loginsCache = null, _loginsCacheTime = 0;
let _distCache   = null, _distCacheTime   = 0;
const CACHE_TTL  = 5 * 60 * 1000;

async function getLogins(cfg) {
  if (_loginsCache && Date.now() - _loginsCacheTime < CACHE_TTL) return _loginsCache;
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_LOGINS || 'Logins'}'!A2:E`,
  });
  _loginsCache = (res.data.values || []).map(r => ({
    firstName: (r[0] || '').trim(),
    lastName:  (r[1] || '').trim(),
    login:     (r[2] || '').trim().toLowerCase(),
    password:  (r[3] || '').trim(),
    tag:       (r[4] || '').trim().replace(/^@/, '').toLowerCase(),
  })).filter(u => u.login);
  _loginsCacheTime = Date.now();
  return _loginsCache;
}

async function getDistribution(cfg) {
  if (_distCache && Date.now() - _distCacheTime < CACHE_TTL) return _distCache;
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_DIST || 'Распределение'}'!A2:B`,
  });
  _distCache = (res.data.values || []).map(r => ({
    object:  (r[0] || '').trim(),
    manager: (r[1] || '').trim(),
  })).filter(d => d.object && d.manager);
  _distCacheTime = Date.now();
  return _distCache;
}

// ── Загрузка непринятых накладных ─────────────────────────────────────────────

async function loadPendingInvoices(cfg) {
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const ID     = process.env.SPREADSHEET_ID;

  const [sentRes, accRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: ID, range: `'${cfg.SHEET_SENT || 'Отправлен'}'!A2:L` }),
    sheets.spreadsheets.values.get({ spreadsheetId: ID, range: `'${cfg.SHEET_ACCEPTED || 'Принят'}'!A2:D` }),
  ]);

  const norm         = s => (s || '').toString().replace(/\s/g, '').toLowerCase();
  const acceptedNums = new Set((accRes.data.values || []).map(r => norm(r[2])).filter(Boolean));
  const todayKey     = todayDateKey();
  const invoices     = [];

  for (const [i, row] of (sentRes.data.values || []).entries()) {
    const obj       = (row[1]  || '').trim();
    const num       = (row[2]  || '').trim();
    const supplier  = (row[4]  || '').trim();
    const dateRaw   = (row[5]  || '').trim();
    const accStatus = (row[10] || '').trim();
    const archive   = (row[11] || '').trim();

    if (!obj || !num) continue;
    if (archive.startsWith('Архив')) continue;
    if (accStatus.startsWith('❌') || accStatus.startsWith('✅')) continue;
    if (acceptedNums.has(norm(num))) continue;

    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;
    const { dd, mm, yyyy } = parsed;
    if (`${yyyy}${mm}${dd}` >= todayKey) continue;

    invoices.push({
      rowIndex: i + 2,
      obj, num, supplier,
      dateStr: `${dd}.${mm}.${yyyy.slice(-2)}`,
      sortKey: `${yyyy}${mm}${dd}`,
    });
  }
  return invoices;
}

// ── Действия над накладной ────────────────────────────────────────────────────

async function applyAction(action, invoice, cfg) {
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const ID     = process.env.SPREADSHEET_ID;
  const SH_SENT = cfg.SHEET_SENT     || 'Отправлен';
  const SH_ACC  = cfg.SHEET_ACCEPTED || 'Принят';

  if (action === 'accept') {
    const now = new Date().toLocaleString('ru-RU', {
      timeZone: 'Europe/Moscow', day: '2-digit', month: '2-digit',
      year: 'numeric', hour: '2-digit', minute: '2-digit',
    });
    const accRes = await sheets.spreadsheets.values.get({ spreadsheetId: ID, range: `'${SH_ACC}'!C:C` });
    const lastRow = (accRes.data.values || []).length + 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId: ID, range: `'${SH_ACC}'!A${lastRow + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[now, invoice.supplier, invoice.num, 'Принят', invoice.dateStr, invoice.obj]] },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: ID, range: `'${SH_SENT}'!K${invoice.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['✅ Принято вручную']] },
    });
  } else {
    const label = action === 'notarrived' ? '❌ Удалено вручную (Не приехала)' : '❌ Удалено вручную';
    await sheets.spreadsheets.values.update({
      spreadsheetId: ID, range: `'${SH_SENT}'!K${invoice.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[label]] },
    });
  }
}

// ── Уведомление в чат ─────────────────────────────────────────────────────────

async function notifyAction(user, obj, action, invoiceNum, cfg) {
  const token    = cfg.TELEGRAM_TOKEN    || process.env.TELEGRAM_TOKEN;
  const chatId   = cfg.TELEGRAM_CHAT_ID  || process.env.TELEGRAM_CHAT_ID;
  const threadId = cfg.TELEGRAM_THREAD_ID || process.env.TELEGRAM_THREAD_ID || null;
  if (!token || !chatId) return;

  const actionLabel = action === 'accept'     ? 'принял(а) накладную'
                    : action === 'notarrived'  ? 'отметил(а) "Не приехала"'
                    : 'удалил(а) накладную';

  await tgPost(token, 'sendMessage', {
    chat_id: chatId,
    text: `${user.firstName} ${user.lastName} ${actionLabel} на ${obj} (${invoiceNum})`,
    ...(threadId ? { message_thread_id: parseInt(threadId) } : {}),
  }).catch(e => console.error('[menuBot] notify error:', e.message));
}

// ── Pending state ─────────────────────────────────────────────────────────────

const pending = new Map();

function newPending(data) {
  const id = Math.random().toString(36).slice(2, 8);
  pending.set(id, data);
  setTimeout(() => pending.delete(id), 60 * 60 * 1000);
  return id;
}

// ── Sessions ──────────────────────────────────────────────────────────────────

const sessions = new Map();

// ── Telegram helpers ──────────────────────────────────────────────────────────

function tgPost(token, method, body) {
  const data = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.telegram.org',
      path:     `/bot${token}/${method}`,
      method:   'POST',
      headers:  { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) },
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch { resolve({}); } });
    });
    req.on('error', reject);
    req.write(data); req.end();
  });
}

function sendMsg(token, chatId, text, keyboard) {
  return tgPost(token, 'sendMessage', {
    chat_id: chatId, text, parse_mode: 'HTML',
    ...(keyboard ? { reply_markup: { inline_keyboard: keyboard } } : {}),
  });
}

function editMsg(token, chatId, messageId, text, keyboard) {
  return tgPost(token, 'editMessageText', {
    chat_id: chatId, message_id: messageId, text, parse_mode: 'HTML',
    reply_markup: { inline_keyboard: keyboard || [] },
  });
}

function answerCbq(token, cbqId, text) {
  return tgPost(token, 'answerCallbackQuery', { callback_query_id: cbqId, text: text || '' });
}

// ── Экраны ────────────────────────────────────────────────────────────────────

async function showAuthScreen(token, chatId, messageId) {
  const text = 'Привет ТУ, авторизуйся:';
  const kbd  = [
    [{ text: '🔑 Авторизация по логину', callback_data: 'mu:authlogin' }],
    [{ text: '👤 Авторизация по тегу',   callback_data: 'mu:authtag'   }],
  ];
  if (messageId) return editMsg(token, chatId, messageId, text, kbd);
  return sendMsg(token, chatId, text, kbd);
}

async function showStatsScreen(token, chatId, messageId, id, state) {
  const { user, objects, invoices } = state;

  const lines = objects.map(obj => {
    const count = invoices.filter(inv => inv.obj === obj).length;
    return `${obj} — ${count} шт.`;
  }).sort();

  const text = `<b>${user.firstName} ${user.lastName}</b>\n\nНепринятые накладные:\n\n${lines.join('\n')}`;
  const kbd  = [[{ text: '📋 Перейти к объектам', callback_data: `mu:objs:${id}:0` }]];

  if (messageId) return editMsg(token, chatId, messageId, text, kbd);
  return sendMsg(token, chatId, text, kbd);
}

const PAGE_SIZE = 10;

async function showObjectsScreen(token, chatId, messageId, id, state, page) {
  const { objects, invoices } = state;
  const total  = objects.length;
  const start  = page * PAGE_SIZE;
  const slice  = objects.slice(start, start + PAGE_SIZE);

  const kbd = slice.map((obj, i) => {
    const count = invoices.filter(inv => inv.obj === obj).length;
    return [{ text: `${obj} (${count} шт.)`, callback_data: `mu:obj:${id}:${start + i}` }];
  });

  const nav = [];
  if (page > 0)                  nav.push({ text: '← Назад',  callback_data: `mu:objs:${id}:${page - 1}` });
  if (start + PAGE_SIZE < total) nav.push({ text: 'Вперёд →', callback_data: `mu:objs:${id}:${page + 1}` });
  if (nav.length) kbd.push(nav);
  kbd.push([{ text: '← К статистике', callback_data: `mu:stats:${id}` }]);

  const text = `<b>Ваши объекты</b> (${start + 1}–${Math.min(start + PAGE_SIZE, total)} из ${total}):`;
  return editMsg(token, chatId, messageId, text, kbd);
}

async function showInvoicesScreen(token, chatId, messageId, id, state, objIdx) {
  const obj     = state.objects[objIdx];
  const invList = state.invoices.filter(inv => inv.obj === obj);

  if (invList.length === 0) {
    return editMsg(token, chatId, messageId, `✅ <b>${obj}</b>\nВсе накладные приняты.`, [
      [{ text: '← Назад', callback_data: `mu:objs:${id}:0` }],
    ]);
  }

  const kbd = invList.map((inv, i) => [{
    text:          `${inv.dateStr} | ${inv.supplier.substring(0, 18)} | ${inv.num}`,
    callback_data: `mu:iv:${id}:${objIdx}:${i}`,
  }]);
  kbd.push([{ text: '← Назад', callback_data: `mu:objs:${id}:0` }]);

  return editMsg(token, chatId, messageId, `<b>${obj}</b>\nВыберите накладную:`, kbd);
}

async function showActionScreen(token, chatId, messageId, id, state, objIdx, invIdx) {
  const obj     = state.objects[objIdx];
  const invList = state.invoices.filter(inv => inv.obj === obj);
  const inv     = invList[invIdx];
  if (!inv) return editMsg(token, chatId, messageId, 'Накладная не найдена', []);

  const kbd = [
    [{ text: '✅ Принять',     callback_data: `mu:ac:${id}:${objIdx}:${invIdx}:accept`     }],
    [{ text: '❌ Удалить',     callback_data: `mu:ac:${id}:${objIdx}:${invIdx}:delete`     }],
    [{ text: '🚫 Не приехала', callback_data: `mu:ac:${id}:${objIdx}:${invIdx}:notarrived` }],
    [{ text: '← Назад',       callback_data: `mu:bk:${id}:inv:${objIdx}`                   }],
  ];

  return editMsg(token, chatId, messageId,
    `<b>Накладная</b>\n🏠 ${obj}\n📅 ${inv.dateStr}\n🏭 ${inv.supplier}\n#️⃣ ${inv.num}`, kbd);
}

// ── Вход после авторизации ────────────────────────────────────────────────────

async function enterMenu(token, chatId, messageId, user, cfg) {
  const [dist, invoices] = await Promise.all([getDistribution(cfg), loadPendingInvoices(cfg)]);

  const fullName = `${user.firstName} ${user.lastName}`.trim().toLowerCase();
  const objects  = dist
    .filter(d => d.manager.trim().toLowerCase() === fullName)
    .map(d => d.object)
    .sort();

  if (objects.length === 0) {
    const text = `✅ ${user.firstName}, вы авторизованы.\nОбъекты не найдены в листе Распределение.`;
    if (messageId) return editMsg(token, chatId, messageId, text, []);
    return sendMsg(token, chatId, text, null);
  }

  const id    = newPending({ user, objects, invoices });
  const state = pending.get(id);
  sessions.set(chatId, { step: 'menu', user });

  await showStatsScreen(token, chatId, messageId, id, state);
}

// ── Публичные обработчики ─────────────────────────────────────────────────────

async function handleStart(msg, token) {
  const chatId = msg.chat.id;
  sessions.set(chatId, { step: 'auth' });
  await showAuthScreen(token, chatId, null);
}

async function handleMenuText(msg, token) {
  const chatId = msg.chat.id;
  const text   = (msg.text || '').trim();
  const sess   = sessions.get(chatId);
  if (!sess || sess.step === 'menu') return false;

  if (sess.step === 'await_login') {
    sessions.set(chatId, { ...sess, step: 'await_password', login: text.toLowerCase() });
    await sendMsg(token, chatId, 'Введите пароль:');
    return true;
  }

  if (sess.step === 'await_password') {
    const cfg   = await getConfig();
    const users = await getLogins(cfg);
    const user  = users.find(u => u.login === sess.login && u.password === text);

    if (!user) {
      sessions.set(chatId, { step: 'auth' });
      await sendMsg(token, chatId, 'Повтори вход или авторизуйся по тегу:', [
        [{ text: '🔑 Авторизация по логину', callback_data: 'mu:authlogin' }],
        [{ text: '👤 Авторизация по тегу',   callback_data: 'mu:authtag'   }],
      ]);
      return true;
    }

    const sent = await sendMsg(token, chatId, '⏳ Загружаю данные...');
    await enterMenu(token, chatId, sent?.result?.message_id, user, cfg);
    return true;
  }

  return false;
}

async function handleMenuCallback(cbq, token) {
  const data      = cbq.data || '';
  const chatId    = cbq.message?.chat?.id;
  const messageId = cbq.message?.message_id;

  if (!data.startsWith('mu:')) return false;

  await answerCbq(token, cbq.id);

  const parts  = data.split(':');
  const action = parts[1];

  try {
    const cfg = await getConfig();

    if (action === 'authlogin') {
      sessions.set(chatId, { step: 'await_login' });
      await editMsg(token, chatId, messageId, 'Введите логин:', []);
      return true;
    }

    if (action === 'authtag') {
      const username = cbq.from?.username?.toLowerCase();
      if (!username) {
        await editMsg(token, chatId, messageId, 'У вас не установлен @тег в Telegram. Обратитесь к Саше.', [
          [{ text: '← Назад', callback_data: 'mu:back_auth' }],
        ]);
        return true;
      }
      const users = await getLogins(cfg);
      const user  = users.find(u => u.tag === username);
      if (!user) {
        await editMsg(token, chatId, messageId, 'Обратитесь к Саше.', [
          [{ text: '← Назад', callback_data: 'mu:back_auth' }],
        ]);
        return true;
      }
      await enterMenu(token, chatId, messageId, user, cfg);
      return true;
    }

    if (action === 'back_auth') {
      sessions.set(chatId, { step: 'auth' });
      await showAuthScreen(token, chatId, messageId);
      return true;
    }

    // Навигация (все остальные действия требуют pending state)
    const id    = parts[2];
    const state = pending.get(id);

    if (!state) {
      await editMsg(token, chatId, messageId, 'Сессия устарела. Введите /start', []);
      return true;
    }

    if (action === 'stats') {
      await showStatsScreen(token, chatId, messageId, id, state);
      return true;
    }

    if (action === 'objs') {
      await showObjectsScreen(token, chatId, messageId, id, state, parseInt(parts[3]) || 0);
      return true;
    }

    if (action === 'obj') {
      await showInvoicesScreen(token, chatId, messageId, id, state, parseInt(parts[3]));
      return true;
    }

    if (action === 'iv') {
      await showActionScreen(token, chatId, messageId, id, state, parseInt(parts[3]), parseInt(parts[4]));
      return true;
    }

    if (action === 'ac') {
      const objIdx  = parseInt(parts[3]);
      const invIdx  = parseInt(parts[4]);
      const actName = parts[5];
      const obj     = state.objects[objIdx];
      const invList = state.invoices.filter(inv => inv.obj === obj);
      const inv     = invList[invIdx];

      if (!inv) {
        await editMsg(token, chatId, messageId, 'Накладная не найдена', []);
        return true;
      }

      await applyAction(actName, inv, cfg);
      await notifyAction(state.user, obj, actName, inv.num, cfg);

      state.invoices = state.invoices.filter(i => i.num !== inv.num);
      pending.set(id, state);

      const remaining = state.invoices.filter(i => i.obj === obj);
      if (remaining.length > 0) {
        await showInvoicesScreen(token, chatId, messageId, id, state, objIdx);
      } else {
        await showObjectsScreen(token, chatId, messageId, id, state, 0);
      }
      return true;
    }

    if (action === 'bk') {
      const target = parts[3];
      if (target === 'inv') {
        await showInvoicesScreen(token, chatId, messageId, id, state, parseInt(parts[4]));
      }
      return true;
    }

  } catch (err) {
    console.error('[menuBot] callback error:', err.message);
    await editMsg(token, chatId, messageId, `Ошибка: ${err.message}`, []);
  }

  return true;
}

module.exports = { handleStart, handleMenuText, handleMenuCallback };
