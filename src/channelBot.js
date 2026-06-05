/**
 * channelBot.js
 *
 * Webhook-сервер для:
 *  1. Приёма сообщений из канала @AcceptODChannel → запись в лист «Принят»
 *  2. Команд от пользователей в личку боту:
 *       /status <объект>  — статистика по конкретному объекту
 *       /status_all       — статистика по всем объектам
 *
 * Колонки «Отправлен» (B=Объект, C=Номер, E=Поставщик, F=Дата, K=Статус приёмки, L=Архив)
 * Колонки «Принят»    (C=Номер, D=Тип, F=Объект)
 */

const express = require('express');
const { getAuthClient, getSheetsClient } = require('./auth');
const { getConfig } = require('./config');
const { ensureSheetExists } = require('./sheets');
const { sendMessage } = require('./telegram');
const { sendErrorsReport } = require('./errorsReport');
const { handleAiMessage, handleCallbackQuery } = require('./aiHandler');

const app = express();
app.use(express.json());

// ── Regex для парсинга сообщений канала ───────────────────────────────────────

const RE_ACCEPTED = /Заказ\s+(#\S+)\s+(.+?)\s+\(поставка\s+(\d{2}-\d{2}-\d{4})\)\s+в\s+ресторане\s+(.+?)\s+был\s+оприходован/i;
const RE_ERROR_A  = /заказа\s+(#\S+)\s+(.+?)\s+в\s+ресторане\s+(.+?)\s+завершилась\s+ошибкой\s+\(поставка\s+(\d{2}-\d{2}-\d{4})\)/i;
const RE_ERROR_B  = /заказа\s+(#\S+)\s+(.+?)\s+\(поставка\s+(\d{2}-\d{2}-\d{4})\)\s+в\s+ресторане\s+(.+?)\s+завершилась\s+ошибкой/i;

function cleanObjectName(name) {
  return (name || '').replace(/\s+(ФГ|ДР|DR|DP|GSW)\s*$/i, '').trim();
}

function parseChannelMessage(text) {
  if (!text) return null;
  const s = text.replace(/\n/g, ' ').replace(/\s{2,}/g, ' ');

  let m = s.match(RE_ACCEPTED);
  if (m) return { type: 'Принят',  orderNumber: m[1].trim(), supplier: m[2].trim(), deliveryDate: m[3].trim(), object: cleanObjectName(m[4]) };

  m = s.match(RE_ERROR_A);
  if (m) return { type: 'Ошибка',  orderNumber: m[1].trim(), supplier: m[2].trim(), object: cleanObjectName(m[3]), deliveryDate: m[4].trim() };

  m = s.match(RE_ERROR_B);
  if (m) return { type: 'Ошибка',  orderNumber: m[1].trim(), supplier: m[2].trim(), deliveryDate: m[3].trim(), object: cleanObjectName(m[4]) };

  return null;
}

// ── Запись в лист «Принят» ────────────────────────────────────────────────────

async function writeToSheet(parsed, rawText, cfg) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SHEET_NAME     = cfg.SHEET_ACCEPTED || 'Принят';

  await ensureSheetExists(SHEET_NAME, [
    'Дата записи', 'Поставщик', 'Номер заказа', 'Тип',
    'Дата поставки', 'Объект', 'Сырое сообщение', '', 'Статус проверки',
  ]);

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const colC = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!C:C`,
  });
  const colValues = colC.data.values || [];
  let lastRow = 1;
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i][0] && colValues[i][0].toString().trim()) { lastRow = i + 1; break; }
  }

  const now = new Date().toLocaleString('ru-RU', {
    timeZone: cfg.TIMEZONE || 'Europe/Moscow',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A${lastRow + 1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[
      now, parsed.supplier, parsed.orderNumber, parsed.type,
      parsed.deliveryDate, parsed.object, rawText.substring(0, 500), '', '',
    ]] },
  });

  console.log(`[channelBot] ✓ ${parsed.type} | ${parsed.orderNumber} | ${parsed.object} | ${parsed.supplier}`);
}

// ── Команды /status и /status_all ─────────────────────────────────────────────

/**
 * Читает листы «Отправлен» и «Принят», возвращает только непринятые строки.
 * Если objectFilter задан — фильтрует по вхождению строки в название объекта.
 *
 * Возвращает массив: [{ supplier, dateStr, dateSortKey, object, status }]
 * status: '⏳' — не оприходовано, '❌' — ошибка регистрации
 */
async function getPendingOrders(cfg, objectFilter = null, supplierFilter = null) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_SENT        = cfg.SHEET_SENT     || 'Отправлен';
  const SH_ACC         = cfg.SHEET_ACCEPTED || 'Принят';

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const [sentRes, accRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_SENT}'!A2:L` }),
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_ACC}'!A2:I` }),
  ]);

  const sentRows = sentRes.data.values || [];
  const accRows  = accRes.data.values  || [];

  const norm = s => (s || '').toString().replace(/\s/g, '').toLowerCase().trim();

  const acceptedNums = new Set();
  const errorNums    = new Set();
  accRows.forEach(r => {
    const num  = norm(r[2]);
    const type = (r[3] || '').toString().trim();
    if (!num) return;
    if (type === 'Принят') acceptedNums.add(num);
    if (type === 'Ошибка') errorNums.add(num);
  });

  const pending = [];

  for (const row of sentRows) {
    const obj      = cleanObjectName((row[1] || '').toString().trim()); // B
    const num      = norm(row[2]);                                       // C
    const supplier = (row[4] || '').toString().trim();                   // E
    const dateRaw  = (row[5] || '').toString().trim();                   // F — Дата отправки
    const archive  = (row[11] || '').toString().trim();                  // L

    if (!obj || !num) continue;
    if (archive.startsWith('Архив')) continue;
    if (acceptedNums.has(num)) continue; // принято — не показываем

    if (objectFilter   && !norm(obj).includes(norm(objectFilter))) continue;
    if (supplierFilter && !norm(supplier).includes(norm(supplierFilter))) continue;

    const accStatus = (row[10] || '').toString().trim();
    if (accStatus.startsWith('❌')) continue;

    const status = errorNums.has(num) ? '❌' : '⏳';

    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;

    const { dd, mm, yyyy } = parsed;
    const { todayDateKey } = require('./dateUtils');
    const sortKey = `${yyyy}${mm}${dd}`;
    if (sortKey >= todayDateKey()) continue;

    const yy = yyyy.slice(-2);
    const dateStr     = `${dd}.${mm}.${yy}`;
    const dateSortKey = sortKey;
    const rawNum = (row[2] || '').toString().trim();

    pending.push({ supplier: supplier || 'Без поставщика', dateStr, dateSortKey, object: obj, status, orderNum: rawNum });
  }

  return pending;
}

/**
 * Строит текст сообщения из списка непринятых строк.
 * Группирует: Поставщик → Дата (новые сверху) → Объекты
 */
function buildPendingText(pending, title) {
  // supplier → dateSortKey → { dateStr, objects[] }
  const bySupplier = {};
  for (const { supplier, dateStr, dateSortKey, object, status } of pending) {
    if (!bySupplier[supplier]) bySupplier[supplier] = {};
    if (!bySupplier[supplier][dateSortKey]) {
      bySupplier[supplier][dateSortKey] = { dateStr, objects: [] };
    }
    bySupplier[supplier][dateSortKey].objects.push({ object, status });
  }

  let text = `*${escMd(title)}*\n`;

  for (const supplier of Object.keys(bySupplier).sort()) {
    text += `\n*${escMd(supplier)}*\n`;

    const dateKeys = Object.keys(bySupplier[supplier]).sort().reverse(); // новые сверху
    for (const dk of dateKeys) {
      const { dateStr, objects } = bySupplier[supplier][dk];
      text += `Дата ${escMd(dateStr)}\n`;
      for (const { object, status } of objects) {
        text += `${status} ${escMd(object)}\n`;
      }
    }
  }

  return text;
}

/**
 * Обрабатывает команду /status <объект>
 */
async function handleStatusCommand(chatId, threadId, objectQuery, cfg) {
  await sendTyping(cfg, chatId, threadId);

  const pending = await getPendingOrders(cfg, objectQuery);

  if (pending.length === 0) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId,
      `✅ Для объекта *${escMd(objectQuery)}* все накладные приняты.`, threadId, 0);
    return;
  }

  const text = buildPendingText(pending, `Не принятые накладные — ${objectQuery}`);
  await sendMessage(cfg.TELEGRAM_TOKEN, chatId, text, threadId, 0);
}

/**
 * Обрабатывает команду /status_all
 */
async function handleStatusAllCommand(chatId, threadId, cfg) {
  await sendTyping(cfg, chatId, threadId);

  const pending = await getPendingOrders(cfg, null);

  if (pending.length === 0) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId,
      '✅ Все накладные приняты.', threadId, 0);
    return;
  }

  // Если текст большой — разбиваем по поставщикам
  const bySupplier = {};
  for (const item of pending) {
    if (!bySupplier[item.supplier]) bySupplier[item.supplier] = [];
    bySupplier[item.supplier].push(item);
  }

  let text = `*Не принятые накладные*\n`;

  for (const supplier of Object.keys(bySupplier).sort()) {
    const chunk = buildPendingText(bySupplier[supplier], supplier).replace(/^\*[^\n]+\n/, '');
    text += `\n*${escMd(supplier)}*\n${chunk}`;

    if (text.length > 3500) {
      await sendMessage(cfg.TELEGRAM_TOKEN, chatId, text, threadId, 0);
      text = '';
    }
  }

  if (text.trim()) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId, text, threadId, 0);
  }
}

// Отправляет "печатает..." в чат
async function sendTyping(cfg, chatId, threadId = null) {
  const https = require('https');
  const typingPayload = { chat_id: chatId, action: 'typing' };
  const tid = parseInt(threadId);
  if (!isNaN(tid) && tid > 0) typingPayload.message_thread_id = tid;
  const body = JSON.stringify(typingPayload);
  return new Promise((resolve) => {
    const req = https.request({
      hostname: 'api.telegram.org',
      path: `/bot${cfg.TELEGRAM_TOKEN}/sendChatAction`,
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    }, res => { res.resume(); res.on('end', resolve); });
    req.on('error', resolve);
    req.write(body); req.end();
  });
}

function escMd(s) {
  return (s || '').toString().replace(/[_*`[]/g, '\\$&');
}

const { parseDateStr } = require('./dateUtils');

// ── Удаление накладной вручную ────────────────────────────────────────────────

async function deleteOrderFromNotifications(chatId, threadId, query, cfg, isTest = false) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_SENT        = cfg.SHEET_SENT || 'Отправлен';
  const norm = s => (s || '').toString().replace(/\s/g, '').toLowerCase().replace(/^#/, '').trim();

  // Парсим: последний токен начинающийся с цифр или # — номер накладной, остальное — объект
  const parts = query.split(/\s+/);
  let orderNum   = '';
  let objectQuery = '';
  for (let i = parts.length - 1; i >= 0; i--) {
    if (/^#?\d/.test(parts[i])) { orderNum = parts[i]; objectQuery = parts.slice(0, i).join(' '); break; }
  }
  if (!orderNum) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId, 'Укажите номер накладной: `/delete Объект #номер`', threadId, 0);
    return;
  }

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_SENT}'!A2:K` });
  const rows   = res.data.values || [];

  const targetNum = norm(orderNum);
  const updates   = [];

  rows.forEach((row, i) => {
    const obj = cleanObjectName((row[1] || '').toString().trim());
    const num = norm(row[2]);
    if (num !== targetNum) return;
    if (objectQuery && !norm(obj).includes(norm(objectQuery))) return;
    if (!isTest) updates.push({ range: `'${SH_SENT}'!K${i + 2}`, values: [['❌ Удалено вручную']] });
  });

  if (updates.length === 0 && !isTest) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId, `Накладная ${escMd(orderNum)} не найдена.`, threadId, 0);
    return;
  }

  if (!isTest) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: updates },
    });
  }
  await sendMessage(cfg.TELEGRAM_TOKEN, chatId,
    isTest ? `Тест: накладная ${escMd(orderNum)} была бы удалена.` : `Накладная ${escMd(orderNum)} убрана из оповещений.`,
    threadId, 0);
}

// ── Вспомогательные функции для статуса ──────────────────────────────────────

function groupBySupplier(pending) {
  const map = {};
  for (const item of pending) {
    if (!map[item.supplier]) map[item.supplier] = {};
    if (!map[item.supplier][item.dateSortKey]) map[item.supplier][item.dateSortKey] = { dateStr: item.dateStr, items: [] };
    map[item.supplier][item.dateSortKey].items.push(item);
  }
  return map;
}

function buildDatesText(dateMap, showObject = true) {
  let text = '';
  for (const key of Object.keys(dateMap).sort()) {
    const { dateStr, items } = dateMap[key];
    text += `  ${dateStr}:\n`;
    for (const item of items) {
      const objPart = showObject ? ` ${escMd(item.object)}` : '';
      text += `    ${item.status}${objPart} ${escMd(item.orderNum || '')}\n`;
    }
  }
  return text;
}

// ── /time_sup ─────────────────────────────────────────────────────────────────

async function handleStatusSupplierCommand(chatId, threadId, supplierQuery, cfg) {
  await sendTyping(cfg, chatId, threadId);
  const pending = await getPendingOrders(cfg, null, supplierQuery);

  if (pending.length === 0) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId, `Все накладные приняты — ${escMd(supplierQuery)}`, threadId, 0);
    return;
  }

  const bySupplier = groupBySupplier(pending);
  let text = `*Не принятые накладные — ${escMd(supplierQuery)}*\n`;
  for (const supplier of Object.keys(bySupplier).sort()) {
    text += `\n*${escMd(supplier)}*\n`;
    text += buildDatesText(bySupplier[supplier], true);
  }
  await sendMessage(cfg.TELEGRAM_TOKEN, chatId, text, threadId, 0);
}

// ── /end_day ──────────────────────────────────────────────────────────────────

async function sendEndOfDayReport(chatId, threadId, cfg) {
  const pending = await getPendingOrders(cfg, null);

  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yy = String(today.getFullYear()).slice(-2);
  const dateLabel = `${dd}.${mm}.${yy}`;

  if (pending.length === 0) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId, `Итог дня ${escMd(dateLabel)}\nВсе накладные приняты.`, threadId, 0);
    return;
  }

  const byObject = {};
  for (const { object } of pending) byObject[object] = (byObject[object] || 0) + 1;

  let text = `*Итог дня ${escMd(dateLabel)}*\nНе принято накладных: ${pending.length}\n\n`;
  for (const obj of Object.keys(byObject).sort()) text += `${escMd(obj)} — ${byObject[obj]}\n`;

  await sendMessage(cfg.TELEGRAM_TOKEN, chatId, text, threadId, 0);
}

// ── /time_all ─────────────────────────────────────────────────────────────────

async function sendTodayOrders(chatId, threadId, cfg) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_SENT        = cfg.SHEET_SENT || 'Отправлен';

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_SENT}'!A2:L` });
  const rows   = res.data.values || [];

  const t = new Date();
  const dd = String(t.getDate()).padStart(2, '0');
  const mm = String(t.getMonth() + 1).padStart(2, '0');
  const yyyy = String(t.getFullYear());
  const todayKey  = `${yyyy}${mm}${dd}`;
  const dateLabel = `${dd}.${mm}.${yyyy.slice(-2)}`;

  const bySupplier = {};
  for (const row of rows) {
    const obj      = cleanObjectName((row[1] || '').toString().trim());
    const supplier = (row[4] || '').toString().trim();
    const dateRaw  = (row[5] || '').toString().trim();
    const archive  = (row[11] || '').toString().trim();
    const accStatus = (row[10] || '').toString().trim();
    if (!obj || !supplier || !dateRaw) continue;
    if (archive.startsWith('Архив')) continue;
    if (accStatus.startsWith('❌')) continue;

    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;
    if (`${parsed.yyyy}${parsed.mm}${parsed.dd}` !== todayKey) continue;

    if (!bySupplier[supplier]) bySupplier[supplier] = [];
    bySupplier[supplier].push(obj);
  }

  const suppliers = Object.keys(bySupplier).sort();
  if (suppliers.length === 0) {
    await sendMessage(cfg.TELEGRAM_TOKEN, chatId, `Сегодня (${dateLabel}) заказов нет.`, threadId, 0);
    return;
  }

  let text = `*Дата заказа ${escMd(dateLabel)}*\nЗаказ отправлен по поставщику:\n`;
  for (const supplier of suppliers) {
    text += `\n*${escMd(supplier)}*\n`;
    for (const obj of bySupplier[supplier]) text += `${escMd(obj)}\n`;
  }
  await sendMessage(cfg.TELEGRAM_TOKEN, chatId, text, threadId, 0);
}

// ── Webhook endpoint ──────────────────────────────────────────────────────────

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);

  try {
    const update = req.body;
    const cfg    = await getConfig();

    // ── Нажатие inline-кнопки (callback_query) ───────────────────────────
    if (update.callback_query) {
      await handleCallbackQuery(update.callback_query, cfg);
      return;
    }

    // ── Личные сообщения / команды ────────────────────────────────────────
    if (update.message) {
      const msg    = update.message;
      const text   = (msg.text || '').trim();
      const chatId = msg.chat.id;
      const replyThreadId = msg.message_thread_id || null;

      // @ReplaceODbot — AI удаление накладной
      if (text.includes('@ReplaceODbot')) {
        await handleAiMessage(msg, cfg);
        return;
      }

      // /delete <объект> <номер>
      if (text.startsWith('/delete_test')) {
        await deleteOrderFromNotifications(chatId, replyThreadId, text.replace(/^\/delete_test\s*/i, '').trim(), cfg, true);
        return;
      }
      if (text.startsWith('/delete')) {
        await deleteOrderFromNotifications(chatId, replyThreadId, text.replace(/^\/delete\s*/i, '').trim(), cfg, false);
        return;
      }

      // /time_sup <поставщик>
      if (text.startsWith('/time_sup')) {
        const query = text.replace(/^\/time_sup\s*/i, '').trim();
        await handleStatusSupplierCommand(chatId, replyThreadId, query, cfg);
        return;
      }

      // /time_all
      if (text === '/time_all' || text.startsWith('/time_all ')) {
        await sendTodayOrders(chatId, replyThreadId, cfg);
        return;
      }

      // /end_day
      if (text === '/end_day') {
        await sendEndOfDayReport(chatId, replyThreadId, cfg);
        return;
      }

      // /errors — Excel с ошибками регистрации
      if (text === '/errors') {
        console.log(`[channelBot] Команда /errors от ${chatId} thread:${replyThreadId}`);
        await sendTyping(cfg, chatId, replyThreadId);
        await sendErrorsReport(chatId, replyThreadId, cfg);
        return;
      }

      // /status_all
      if (text === '/status_all' || text.startsWith('/status_all ')) {
        console.log(`[channelBot] Команда /status_all от ${chatId} thread:${replyThreadId}`);
        await handleStatusAllCommand(chatId, replyThreadId, cfg);
        return;
      }

      // /status <объект>
      if (text.startsWith('/status')) {
        const query = text.replace(/^\/status\s*/i, '').trim();
        if (!query) {
          await sendMessage(cfg.TELEGRAM_TOKEN, chatId,
            'Укажите объект: `/status OD Нахабино`\nИли все: `/status_all`', replyThreadId, 0);
          return;
        }
        console.log(`[channelBot] Команда /status "${query}" от ${chatId} thread:${replyThreadId}`);
        await handleStatusCommand(chatId, replyThreadId, query, cfg);
        return;
      }
    }

    // ── Сообщения из канала ───────────────────────────────────────────────
    const msg = update.channel_post;
    if (!msg) return;

    const text = msg.text || msg.caption || '';
    if (!text) return;

    const parsed = parseChannelMessage(text);
    if (!parsed) {
      console.log(`[channelBot] Пропущено: ${text.substring(0, 80)}`);
      return;
    }

    await writeToSheet(parsed, text, cfg);

  } catch (err) {
    console.error(`[channelBot] Ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
});

app.get('/health', (req, res) =>
  res.json({ status: 'ok', ts: new Date().toISOString() }));

// ── Регистрация webhook ───────────────────────────────────────────────────────

async function registerWebhook(token, webhookUrl) {
  const https = require('https');
  // Разрешаем и channel_post и message (для команд)
  const body = JSON.stringify({
    url: webhookUrl,
    allowed_updates: ['channel_post', 'message', 'callback_query'],
  });

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.telegram.org',
      path: `/bot${token}/setWebhook`,
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        const r = JSON.parse(data);
        if (r.ok) console.log(`[channelBot] Webhook зарегистрирован: ${webhookUrl}`);
        else      console.error(`[channelBot] Ошибка webhook: ${r.description}`);
        resolve(r);
      });
    });
    req.on('error', reject);
    req.write(body); req.end();
  });
}

async function startChannelBot() {
  const cfg        = await getConfig();
  const token      = cfg.TELEGRAM_BOT_CHANNEL_TOKEN || cfg.TELEGRAM_TOKEN;
  const webhookUrl = process.env.WEBHOOK_URL;
  const port       = process.env.PORT || 3000;

  app.listen(port, () =>
    console.log(`[channelBot] HTTP-сервер на порту ${port}`));

  if (webhookUrl && token) {
    await registerWebhook(token, `${webhookUrl}/webhook`);
  } else {
    console.warn('[channelBot] WEBHOOK_URL не задан — добавьте в Railway Variables');
  }
}

module.exports = { startChannelBot, parseChannelMessage, sendTodayOrders, sendEndOfDayReport };
