/**
 * menuBot.js — интерактивное меню бота с авторизацией.
 *
 * Флоу:
 *  /start → запрос логина → запрос пароля → меню дат
 *  Дата → список объектов → список накладных → действие
 *
 * Действия: Принять / Удалить / Не приехала / Назад
 */

const https = require('https');
const { getAuthClient, getSheetsClient } = require('./auth');
const { getConfig } = require('./config');
const { parseDateStr, todayDateKey } = require('./dateUtils');

// ── Авторизация ───────────────────────────────────────────────────────────────

const AUTH_LOGIN    = 'alexberv';
const AUTH_PASSWORD = '7290';

// chatId → { step: 'login'|'password'|'menu' }
const sessions = new Map();

function getSession(chatId) {
  return sessions.get(chatId) || null;
}

function setSession(chatId, data) {
  sessions.set(chatId, data);
}

function isAuthed(chatId) {
  const s = sessions.get(chatId);
  return s && s.step === 'menu';
}

// ── Pending state для навигации ───────────────────────────────────────────────

// id → { level, dates?, selDate?, objects?, selObj?, invoices? }
const pending = new Map();

function newPending(data) {
  const id = Math.random().toString(36).slice(2, 8);
  pending.set(id, data);
  setTimeout(() => pending.delete(id), 30 * 60 * 1000); // TTL 30 мин
  return id;
}

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
    req.write(data);
    req.end();
  });
}

function sendMsg(token, chatId, text, keyboard) {
  return tgPost(token, 'sendMessage', {
    chat_id:      chatId,
    text,
    parse_mode:   'HTML',
    ...(keyboard ? { reply_markup: { inline_keyboard: keyboard } } : {}),
  });
}

function editMsg(token, chatId, messageId, text, keyboard) {
  return tgPost(token, 'editMessageText', {
    chat_id:      chatId,
    message_id:   messageId,
    text,
    parse_mode:   'HTML',
    ...(keyboard ? { reply_markup: { inline_keyboard: keyboard } } : { reply_markup: { inline_keyboard: [] } }),
  });
}

function answerCbq(token, cbqId, text) {
  return tgPost(token, 'answerCallbackQuery', { callback_query_id: cbqId, text: text || '' });
}

// ── Загрузка непринятых накладных ─────────────────────────────────────────────

async function loadPendingInvoices(cfg) {
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_SENT = cfg.SHEET_SENT     || 'Отправлен';
  const SH_ACC  = cfg.SHEET_ACCEPTED || 'Принят';

  const [sentRes, accRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_SENT}'!A2:L` }),
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_ACC}'!A2:D` }),
  ]);

  const norm = s => (s || '').toString().replace(/\s/g, '').toLowerCase();
  const acceptedNums = new Set((accRes.data.values || []).map(r => norm(r[2])).filter(Boolean));

  const todayKey = todayDateKey();
  const invoices = [];

  for (const [i, row] of (sentRes.data.values || []).entries()) {
    const obj       = (row[1] || '').toString().trim();
    const num       = (row[2] || '').toString().trim();
    const supplier  = (row[4] || '').toString().trim();
    const dateRaw   = (row[5] || '').toString().trim();
    const accStatus = (row[10] || '').toString().trim();
    const archive   = (row[11] || '').toString().trim();

    if (!obj || !num) continue;
    if (archive.startsWith('Архив')) continue;
    if (accStatus.startsWith('❌') || accStatus.startsWith('✅')) continue;
    if (acceptedNums.has(norm(num))) continue;

    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;
    const { dd, mm, yyyy } = parsed;
    const sortKey = `${yyyy}${mm}${dd}`;
    if (sortKey >= todayKey) continue;

    invoices.push({
      rowIndex: i + 2, // 1-based row in sheet (A2 = row 2)
      obj, num, supplier,
      dateStr:  `${dd}.${mm}.${yyyy.slice(-2)}`,
      sortKey,
    });
  }

  return invoices;
}

// ── Действия над накладной ────────────────────────────────────────────────────

async function applyAction(action, invoice, cfg) {
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_SENT = cfg.SHEET_SENT     || 'Отправлен';
  const SH_ACC  = cfg.SHEET_ACCEPTED || 'Принят';

  if (action === 'accept') {
    // Пишем в лист Принят
    const now = new Date().toLocaleString('ru-RU', {
      timeZone: 'Europe/Moscow',
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit',
    });
    const accRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SH_ACC}'!C:C`,
    });
    const lastRow = (accRes.data.values || []).length + 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SH_ACC}'!A${lastRow + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[now, invoice.supplier, invoice.num, 'Принят', invoice.dateStr, invoice.obj]] },
    });
    // Ставим статус в Отправлен K
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SH_SENT}'!K${invoice.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['✅ Принято вручную']] },
    });

  } else {
    const label = action === 'delete'    ? '❌ Удалено вручную'
                : action === 'notarrived'? '❌ Удалено вручную (Не приехала)'
                : '❌ Удалено вручную';
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SH_SENT}'!K${invoice.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[label]] },
    });
  }
}

// ── Построение меню ───────────────────────────────────────────────────────────

function buildDatesMenu(invoices) {
  // Уникальные даты, отсортированные
  const dates = [...new Set(invoices.map(inv => inv.sortKey))]
    .sort()
    .map(key => {
      const inv = invoices.find(i => i.sortKey === key);
      const count = invoices.filter(i => i.sortKey === key).length;
      return { sortKey: key, dateStr: inv.dateStr, count };
    });
  return dates;
}

function buildObjectsMenu(invoices, sortKey) {
  const filtered = invoices.filter(i => i.sortKey === sortKey);
  const objects = [...new Set(filtered.map(i => i.obj))].sort();
  return objects.map(obj => ({
    obj,
    count: filtered.filter(i => i.obj === obj).length,
  }));
}

function buildInvoicesList(invoices, sortKey, obj) {
  return invoices.filter(i => i.sortKey === sortKey && i.obj === obj);
}

// ── Отправка экранов ──────────────────────────────────────────────────────────

async function showDatesScreen(token, chatId, messageId, invoices) {
  const dates = buildDatesMenu(invoices);

  if (dates.length === 0) {
    const text = '✅ Все накладные приняты';
    if (messageId) return editMsg(token, chatId, messageId, text, null);
    return sendMsg(token, chatId, text, null);
  }

  const id  = newPending({ level: 'dates', invoices });
  const kbd = dates.map(d => [{
    text:          `📅 ${d.dateStr}  (${d.count} шт.)`,
    callback_data: `m:dt:${id}:${d.sortKey}`,
  }]);

  const text = `<b>Непринятые накладные</b>\nВыберите дату:`;
  if (messageId) return editMsg(token, chatId, messageId, text, kbd);
  return sendMsg(token, chatId, text, kbd);
}

async function showObjectsScreen(token, chatId, messageId, id, state, sortKey) {
  const objects = buildObjectsMenu(state.invoices, sortKey);
  const dateStr = state.invoices.find(i => i.sortKey === sortKey)?.dateStr || sortKey;

  const kbd = objects.map(o => [{
    text:          `🏠 ${o.obj}  (${o.count} шт.)`,
    callback_data: `m:ob:${id}:${sortKey}:${encodeObjIdx(state.invoices, sortKey, o.obj)}`,
  }]);
  kbd.push([{ text: '← Назад', callback_data: `m:bk:${id}:dates` }]);

  return editMsg(token, chatId, messageId,
    `<b>Дата: ${dateStr}</b>\nВыберите объект:`, kbd);
}

async function showInvoicesScreen(token, chatId, messageId, id, state, sortKey, objIdx) {
  const objects = buildObjectsMenu(state.invoices, sortKey);
  const obj     = objects[objIdx]?.obj;
  if (!obj) return editMsg(token, chatId, messageId, 'Ошибка: объект не найден', null);

  const invList = buildInvoicesList(state.invoices, sortKey, obj);
  const dateStr = state.invoices.find(i => i.sortKey === sortKey)?.dateStr || sortKey;

  const kbd = invList.map((inv, i) => [{
    text:          `${inv.dateStr} | ${inv.supplier.substring(0, 20)} | ${inv.num}`,
    callback_data: `m:iv:${id}:${sortKey}:${objIdx}:${i}`,
  }]);
  kbd.push([{ text: '← Назад', callback_data: `m:bk:${id}:obj:${sortKey}` }]);

  return editMsg(token, chatId, messageId,
    `<b>${dateStr} / ${obj}</b>\nВыберите накладную:`, kbd);
}

async function showActionScreen(token, chatId, messageId, id, state, sortKey, objIdx, invIdx) {
  const objects = buildObjectsMenu(state.invoices, sortKey);
  const obj     = objects[objIdx]?.obj;
  const invList = buildInvoicesList(state.invoices, sortKey, obj);
  const inv     = invList[invIdx];
  if (!inv) return editMsg(token, chatId, messageId, 'Ошибка: накладная не найдена', null);

  const kbd = [
    [{ text: '✅ Принять',      callback_data: `m:ac:${id}:${sortKey}:${objIdx}:${invIdx}:accept` }],
    [{ text: '❌ Удалить',      callback_data: `m:ac:${id}:${sortKey}:${objIdx}:${invIdx}:delete` }],
    [{ text: '🚫 Не приехала',  callback_data: `m:ac:${id}:${sortKey}:${objIdx}:${invIdx}:notarrived` }],
    [{ text: '← Назад',        callback_data: `m:bk:${id}:inv:${sortKey}:${objIdx}` }],
  ];

  return editMsg(token, chatId, messageId,
    `<b>Накладная</b>\n📅 ${inv.dateStr}\n🏠 ${obj}\n🏭 ${inv.supplier}\n#️⃣ ${inv.num}`,
    kbd);
}

// Вспомогательная: индекс объекта в списке объектов для dateKey
function encodeObjIdx(invoices, sortKey, obj) {
  const objects = buildObjectsMenu(invoices, sortKey);
  return objects.findIndex(o => o.obj === obj);
}

// ── Публичные обработчики ─────────────────────────────────────────────────────

async function handleStart(msg, token) {
  const chatId = msg.chat.id;
  setSession(chatId, { step: 'login' });
  await sendMsg(token, chatId, 'Введите логин:');
}

async function handleMenuText(msg, token) {
  const chatId = msg.chat.id;
  const text   = (msg.text || '').trim();
  const sess   = getSession(chatId);

  if (!sess) return; // не в процессе авторизации

  if (sess.step === 'login') {
    if (text === AUTH_LOGIN) {
      setSession(chatId, { step: 'password' });
      await sendMsg(token, chatId, 'Введите пароль:');
    } else {
      await sendMsg(token, chatId, '❌ Неверный логин. Попробуйте /start');
      sessions.delete(chatId);
    }
    return true;
  }

  if (sess.step === 'password') {
    if (text === AUTH_PASSWORD) {
      setSession(chatId, { step: 'menu' });
      const cfg      = await getConfig();
      const invoices = await loadPendingInvoices(cfg);
      await showDatesScreen(token, chatId, null, invoices);
    } else {
      await sendMsg(token, chatId, '❌ Неверный пароль. Попробуйте /start');
      sessions.delete(chatId);
    }
    return true;
  }

  return false; // не наш — пусть обрабатывает основной handler
}

async function handleMenuCallback(cbq, token) {
  const data      = cbq.data || '';
  const chatId    = cbq.message?.chat?.id;
  const messageId = cbq.message?.message_id;

  if (!data.startsWith('m:')) return false;

  await answerCbq(token, cbq.id);

  const parts = data.split(':');
  // parts[0] = 'm', parts[1] = action, parts[2] = id, parts[3+] = args

  const action = parts[1];
  const id     = parts[2];
  const state  = pending.get(id);

  if (!state) {
    await editMsg(token, chatId, messageId, 'Сессия устарела. Введите /start', null);
    return true;
  }

  try {
    const cfg = await getConfig();

    if (action === 'dt') {
      // Выбор даты → показ объектов
      const sortKey = parts[3];
      await showObjectsScreen(token, chatId, messageId, id, state, sortKey);

    } else if (action === 'ob') {
      // Выбор объекта → показ накладных
      const sortKey = parts[3];
      const objIdx  = parseInt(parts[4]);
      await showInvoicesScreen(token, chatId, messageId, id, state, sortKey, objIdx);

    } else if (action === 'iv') {
      // Выбор накладной → показ действий
      const sortKey = parts[3];
      const objIdx  = parseInt(parts[4]);
      const invIdx  = parseInt(parts[5]);
      await showActionScreen(token, chatId, messageId, id, state, sortKey, objIdx, invIdx);

    } else if (action === 'ac') {
      // Выполнение действия
      const sortKey   = parts[3];
      const objIdx    = parseInt(parts[4]);
      const invIdx    = parseInt(parts[5]);
      const actName   = parts[6];

      const objects = buildObjectsMenu(state.invoices, sortKey);
      const obj     = objects[objIdx]?.obj;
      const invList = buildInvoicesList(state.invoices, sortKey, obj);
      const inv     = invList[invIdx];

      if (!inv) {
        await editMsg(token, chatId, messageId, 'Накладная не найдена', null);
        return true;
      }

      await applyAction(actName, inv, cfg);

      const label = actName === 'accept'     ? '✅ Принято'
                  : actName === 'delete'      ? '❌ Удалено'
                  : actName === 'notarrived'  ? '🚫 Не приехала'
                  : 'Готово';

      // Убираем обработанную накладную из state и обновляем pending
      state.invoices = state.invoices.filter(i => i.num !== inv.num);
      pending.set(id, state);

      // Показываем подтверждение и возвращаемся к списку накладных
      const remaining = buildInvoicesList(state.invoices, sortKey, obj);
      if (remaining.length > 0) {
        await editMsg(token, chatId, messageId,
          `${label}: ${inv.num}\n\nВозвращаемся к списку накладных...`, null);
        await showInvoicesScreen(token, chatId, messageId, id, state, sortKey, objIdx);
      } else {
        // Накладных по этому объекту не осталось — возврат к объектам
        await editMsg(token, chatId, messageId,
          `${label}: ${inv.num}\n\nВсе накладные объекта обработаны.`, null);
        const objsLeft = buildObjectsMenu(state.invoices, sortKey);
        if (objsLeft.length > 0) {
          await showObjectsScreen(token, chatId, messageId, id, state, sortKey);
        } else {
          await showDatesScreen(token, chatId, messageId, state.invoices);
        }
      }

    } else if (action === 'bk') {
      // Назад
      const target = parts[3];
      if (target === 'dates') {
        await showDatesScreen(token, chatId, messageId, state.invoices);
      } else if (target === 'obj') {
        const sortKey = parts[4];
        await showObjectsScreen(token, chatId, messageId, id, state, sortKey);
      } else if (target === 'inv') {
        const sortKey = parts[4];
        const objIdx  = parseInt(parts[5]);
        await showInvoicesScreen(token, chatId, messageId, id, state, sortKey, objIdx);
      }
    }
  } catch (err) {
    console.error('[menuBot] handleMenuCallback:', err.message);
    await editMsg(token, chatId, messageId, `Ошибка: ${err.message}`, null);
  }

  return true;
}

module.exports = { handleStart, handleMenuText, handleMenuCallback, isAuthed };
