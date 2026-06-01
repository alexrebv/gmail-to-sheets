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
 * Парсит дату из строки без new Date() (избегаем UTC-сдвига).
 * Форматы: "2026-05-31 11:50:00", "31.05.2026, 11:50", "31.05.2026"
 * Возвращает { dd, mm, yyyy } или null.
 */
function parseDateStr(s) {
  if (!s) return null;
  s = s.toString().trim();

  // ISO-like: 2026-05-31 или 2026-05-31 11:50:00
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return { yyyy: iso[1], mm: iso[2], dd: iso[3] };

  // Русский: 31.05.2026 или 31.05.2026, 11:50 или 31.05.26
  const ru = s.match(/^(\d{2})\.(\d{2})\.(\d{2,4})/);
  if (ru) {
    const yyyy = ru[3].length === 2 ? `20${ru[3]}` : ru[3];
    return { yyyy, mm: ru[2], dd: ru[1] };
  }

  return null;
}

/** Возвращает YYYYMMDD сегодняшней даты в локальном времени */
function todayDateKey() {
  const t = new Date();
  const dd   = String(t.getDate()).padStart(2, '0');
  const mm   = String(t.getMonth() + 1).padStart(2, '0');
  const yyyy = String(t.getFullYear());
  return `${yyyy}${mm}${dd}`;
}

/**
 * Читает листы «Отправлен» и «Принят», возвращает только непринятые строки.
 * Если objectFilter задан — фильтрует по вхождению строки в название объекта.
 *
 * Возвращает массив: [{ supplier, dateStr, dateSortKey, object, status }]
 * status: '⏳' — не оприходовано, '❌' — ошибка регистрации
 */
async function getPendingOrders(cfg, objectFilter = null) {
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

    if (objectFilter && !norm(obj).includes(norm(objectFilter))) continue;

    const status = errorNums.has(num) ? '❌' : '⏳';

    // Парсим дату строкой — без new Date() во избежание UTC-сдвига
    // Поддерживаемые форматы:
    //   2026-05-31 11:50:00   (ISO-like)
    //   31.05.2026, 11:50     (русский с запятой)
    //   31.05.2026            (русский без времени)
    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;

    const { dd, mm, yyyy } = parsed;
    const todayKey = todayDateKey();
    const sortKey  = `${yyyy}${mm}${dd}`;
    if (sortKey >= todayKey) continue; // сегодняшние — не показываем

    const yy = yyyy.slice(-2);
    const dateStr     = `${dd}.${mm}.${yy}`;
    const dateSortKey = sortKey;

    pending.push({ supplier: supplier || 'Без поставщика', dateStr, dateSortKey, object: obj, status });
  }

  return pending;
}

/** Форматирует блок дат → объекты (без эмодзи, без названия поставщика) */
function buildDatesText(dateMap) {
  let text = '';
  const dateKeys = Object.keys(dateMap).sort().reverse(); // новые сверху
  for (const dk of dateKeys) {
    const { dateStr, objects } = dateMap[dk];
    text += `Дата ${escMd(dateStr)}\n`;
    for (const { object } of objects) {
      text += `${escMd(object)}\n`;
    }
  }
  return text;
}

/** Группирует pending по поставщику → dateMap */
function groupBySupplier(pending) {
  const bySupplier = {};
  for (const { supplier, dateStr, dateSortKey, object } of pending) {
    if (!bySupplier[supplier]) bySupplier[supplier] = {};
    if (!bySupplier[supplier][dateSortKey]) {
      bySupplier[supplier][dateSortKey] = { dateStr, objects: [] };
    }
    bySupplier[supplier][dateSortKey].objects.push({ object });
  }
  return bySupplier;
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

  const bySupplier = groupBySupplier(pending);
  let text = `*Не принятые накладные — ${escMd(objectQuery)}*\n`;

  for (const supplier of Object.keys(bySupplier).sort()) {
    text += `\n*${escMd(supplier)}*\n`;
    text += buildDatesText(bySupplier[supplier]);
  }

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

  const bySupplier = groupBySupplier(pending);
  let text = `*Не принятые накладные*\n`;

  for (const supplier of Object.keys(bySupplier).sort()) {
    text += `\n*${escMd(supplier)}*\n`;
    text += buildDatesText(bySupplier[supplier]);

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

// ── Webhook endpoint ──────────────────────────────────────────────────────────

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);

  try {
    const update = req.body;
    const cfg    = await getConfig();

    // ── Личные сообщения / команды ────────────────────────────────────────
    if (update.message) {
      const msg    = update.message;
      const text   = (msg.text || '').trim();
      const chatId = msg.chat.id;
      // Берём thread_id из входящего сообщения — отвечаем в тот же топик
      const replyThreadId = msg.message_thread_id || null;

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
    allowed_updates: ['channel_post', 'message'],
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

module.exports = { startChannelBot, parseChannelMessage };
