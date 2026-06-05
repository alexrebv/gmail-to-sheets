/**
 * aiHandler.js
 *
 * Обрабатывает команды через @ReplaceODbot с распознаванием речи (Claude Haiku).
 *
 * Флоу:
 *  1. Пользователь пишет "@ReplaceODbot удали курский 07, не приехала"
 *  2. Claude извлекает: action, objectHint, orderNum, reason (2 слова)
 *  3. Если объект неоднозначен → инлайн кнопки с вариантами
 *  4. Если накладная не указана → инлайн кнопки с непринятыми накладными объекта
 *  5. Удаляет: ставит K = "❌ Удалено вручную (причина)"
 */

const https  = require('https');
const Anthropic = require('@anthropic-ai/sdk');
const { getAuthClient, getSheetsClient } = require('./auth');
const { parseDateStr, todayDateKey }     = require('./dateUtils');

// ── Кэш объектов ─────────────────────────────────────────────────────────────

let _objectsCache    = null;
let _objectsCacheTime = 0;
const OBJECTS_TTL    = 10 * 60 * 1000;

async function getObjects(cfg) {
  if (_objectsCache && Date.now() - _objectsCacheTime < OBJECTS_TTL) return _objectsCache;

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_OBJECTS || 'Объекты'}'!A2:A`,
  });

  _objectsCache    = (res.data.values || []).map(r => (r[0] || '').toString().trim()).filter(Boolean);
  _objectsCacheTime = Date.now();
  return _objectsCache;
}

// ── Claude Haiku — парсинг intent ─────────────────────────────────────────────

const SYSTEM_PROMPT = `Из сообщения пользователя извлеки поля. Отвечай ТОЛЬКО компактным JSON без пробелов и переносов.
Формат: {"action":"delete"|"unknown","objectHint":string|null,"supplierHint":string|null,"dateHint":string|null,"orderNum":string|null,"reason":string|null}
action: "delete" если в сообщении речь идёт об удалении/отмене/случайном удалении заказа/накладной — в любом времени и форме (удали, удалили, отменили, случайно удалили, убери, убрать, не нужна, не приехала, ошибочно)
objectHint: название объекта/ресторана из сообщения или null. "на Одинцово" → "Одинцово"
supplierHint: название поставщика или null
dateHint: дата в формате DD.MM (только день и месяц, например "04.06", "03.06") или null
orderNum: номер накладной (формат #20260-xxx-xxxx или цифры с дефисами) или null
reason: ПРИЧИНА в 2 слова на русском отражающая суть. Примеры: "Случайно удалили", "Не приехала", "Уже принята", "Ошибка заказа", "Дубль накладной", "Не заказывали". Если причина не указана но ясна из контекста — выведи её сам`;

async function parseIntent(text) {
  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const msg = await client.messages.create({
    model:      'claude-haiku-4-5-20251001',
    max_tokens: 120,
    system:     SYSTEM_PROMPT,
    messages:   [{ role: 'user', content: text }],
  });

  try {
    return JSON.parse(msg.content[0].text);
  } catch {
    return { action: 'unknown' };
  }
}

// ── Поиск объектов по подстроке ───────────────────────────────────────────────

function findMatchingObjects(hint, objects) {
  if (!hint) return [];

  const norm = s => s.toLowerCase().replace(/[-\s]/g, '');
  const h    = norm(hint);

  // Разбиваем хинт на буквенную и числовую части
  // "Курс07" → letters="курс", digits="07"
  // "Курский 07" → letters="курский", digits="07"
  const letters = h.replace(/\d/g, '').trim();
  const digits  = h.replace(/\D/g, '').trim();

  return objects.filter(obj => {
    const o = norm(obj);
    // Сначала точное вхождение хинта
    if (o.includes(h)) return true;
    // Затем: буквы И цифры по отдельности
    const lettersOk = !letters || o.includes(letters);
    const digitsOk  = !digits  || o.includes(digits);
    return lettersOk && digitsOk;
  });
}

// ── Непринятые накладные по объекту ──────────────────────────────────────────

async function getPendingInvoicesForObject(objectName, cfg, supplierHint = null, dateHint = null) {
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_SENT || 'Отправлен'}'!A2:L`,
  });

  const rows     = res.data.values || [];
  const todayKey = todayDateKey();
  const normStr  = s => (s || '').toString().replace(/[-\s]/g, '').toLowerCase().trim();

  // dateHint: "04.06" → dd="04", mm="06"
  let filterDd = null, filterMm = null;
  if (dateHint) {
    const dm = dateHint.match(/^(\d{1,2})\.(\d{1,2})$/);
    if (dm) { filterDd = dm[1].padStart(2, '0'); filterMm = dm[2].padStart(2, '0'); }
  }

  const invoices = [];
  for (const row of rows) {
    const obj       = (row[1] || '').toString().trim();
    const numRaw    = (row[2] || '').toString().trim();
    const supplier  = (row[4] || '').toString().trim();
    const dateRaw   = (row[5] || '').toString().trim();
    const archive   = (row[11] || '').toString().trim();
    const accStatus = (row[10] || '').toString().trim();

    if (!obj || !numRaw || !dateRaw) continue;
    if (archive.startsWith('Архив')) continue;
    if (accStatus.startsWith('❌') || accStatus.startsWith('✅')) continue;
    if (normStr(obj) !== normStr(objectName)) continue;

    // Фильтр по поставщику
    if (supplierHint && !normStr(supplier).includes(normStr(supplierHint))) continue;

    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;
    const { dd, mm, yyyy } = parsed;
    if (`${yyyy}${mm}${dd}` >= todayKey) continue;

    // Фильтр по дате
    if (filterDd && filterMm && (dd !== filterDd || mm !== filterMm)) continue;

    invoices.push({ num: numRaw, dateStr: `${dd}.${mm}.${yyyy.slice(-2)}`, supplier });
  }

  return invoices;
}

// ── Удалить накладную ─────────────────────────────────────────────────────────

async function markDeleted(orderNum, reason, cfg) {
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const SH     = cfg.SHEET_SENT || 'Отправлен';

  const res  = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${SH}'!A2:K`,
  });
  const rows = res.data.values || [];

  const normNum = s => (s || '').toString().replace(/\s/g, '').replace(/^#/, '').toLowerCase();
  const target  = normNum(orderNum);
  const updates = [];

  rows.forEach((row, i) => {
    if (normNum(row[2]) === target) {
      const label = reason ? `❌ Удалено вручную (${reason})` : '❌ Удалено вручную';
      updates.push({ range: `'${SH}'!K${i + 2}`, values: [[label]] });
    }
  });

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: process.env.SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: updates },
    });
    console.log(`[aiHandler] Удалено: ${orderNum} — ${reason}`);
    return true;
  }
  return false;
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
      res.on('end', () => resolve(JSON.parse(d)));
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function sendKeyboard(token, chatId, threadId, text, keyboard) {
  return tgPost(token, 'sendMessage', {
    chat_id:      chatId,
    text,
    reply_markup: { inline_keyboard: keyboard },
    ...(threadId ? { message_thread_id: parseInt(threadId) } : {}),
  });
}

async function editKeyboard(token, chatId, messageId, text) {
  return tgPost(token, 'editMessageText', {
    chat_id:    chatId,
    message_id: messageId,
    text,
  });
}

async function answerCbq(token, cbqId, text) {
  return tgPost(token, 'answerCallbackQuery', { callback_query_id: cbqId, text });
}

// ── Pending state ─────────────────────────────────────────────────────────────

// id → { step, objects?, invoices?, objectName?, orderNum?, reason, chatId, threadId }
const pending = new Map();

function newPending(data) {
  const id = Math.random().toString(36).slice(2, 8);
  pending.set(id, data);
  setTimeout(() => pending.delete(id), 5 * 60 * 1000); // TTL 5 мин
  return id;
}

// ── Основной обработчик входящего сообщения ────────────────────────────────────

async function handleAiMessage(msg, cfg) {
  const rawText  = (msg.text || msg.caption || '').trim();
  const chatId   = msg.chat.id;
  const threadId = msg.message_thread_id || null;
  const token    = cfg.TELEGRAM_TOKEN;

  const { sendMessage } = require('./telegram');
  const reply = (t) => sendMessage(token, chatId, t, threadId, 0);

  // Убираем упоминание бота
  const text = rawText.replace(/@\w+/g, '').trim();

  try {
    const objects = await getObjects(cfg);
    const intent  = await parseIntent(text);

    if (intent.action !== 'delete') {
      return reply('Повторите запрос');
    }

    const reason       = intent.reason || null;
    const orderNum     = intent.orderNum ? intent.orderNum.replace(/^#/, '') : null;
    const supplierHint = intent.supplierHint || null;
    const dateHint     = intent.dateHint || null;

    const matches = findMatchingObjects(intent.objectHint, objects);

    if (matches.length === 0) {
      return reply('Не распознал объект');
    }

    if (matches.length > 10) {
      return reply('Уточните название объекта — слишком много совпадений');
    }

    if (matches.length === 1) {
      return proceedWithObject(matches[0], orderNum, reason, chatId, threadId, token, cfg, supplierHint, dateHint);
    }

    // Несколько объектов — кнопки выбора
    const id  = newPending({ step: 'select_object', objects: matches, orderNum, reason, supplierHint, dateHint });
    const kbd = matches.map((obj, idx) => [{ text: obj, callback_data: `po:${id}:${idx}` }]);
    await sendKeyboard(token, chatId, threadId, 'Уточните объект:', kbd);

  } catch (err) {
    console.error('[aiHandler] handleAiMessage:', err.message);
    await reply('Повторите запрос');
  }
}

async function proceedWithObject(objectName, orderNum, reason, chatId, threadId, token, cfg, supplierHint = null, dateHint = null) {
  const { sendMessage } = require('./telegram');
  const reply = (t) => sendMessage(token, chatId, t, threadId, 0);

  if (orderNum) {
    const ok = await markDeleted(orderNum, reason, cfg);
    return reply(ok ? 'Принято' : 'Накладная не найдена');
  }

  // Нет номера — ищем непринятые по объекту (+ фильтры по поставщику и дате)
  const invoices = await getPendingInvoicesForObject(objectName, cfg, supplierHint, dateHint);

  if (invoices.length === 0) return reply(`Нет непринятых накладных — ${objectName}`);
  if (invoices.length === 1) {
    const ok = await markDeleted(invoices[0].num, reason, cfg);
    return reply(ok ? 'Принято' : 'Накладная не найдена');
  }

  // Несколько накладных — кнопки
  const id  = newPending({ step: 'select_invoice', invoices, objectName, reason });
  const kbd = invoices.map((inv, idx) => [{
    text:          `${inv.num} (${inv.dateStr})`,
    callback_data: `pi:${id}:${idx}`,
  }]);

  await sendKeyboard(token, chatId, threadId, `Выберите накладную — ${objectName}:`, kbd);
}

// ── Обработчик нажатия кнопки (callback_query) ───────────────────────────────

async function handleCallbackQuery(cbq, cfg) {
  const token     = cfg.TELEGRAM_TOKEN;
  const data      = cbq.data || '';
  const chatId    = cbq.message?.chat?.id;
  const messageId = cbq.message?.message_id;
  const threadId  = cbq.message?.message_thread_id || null;

  const { sendMessage } = require('./telegram');
  const reply = (t) => sendMessage(token, chatId, t, threadId, 0);

  const parts = data.split(':');
  const type  = parts[0];
  const id    = parts[1];
  const idx   = parseInt(parts[2]);

  const state = pending.get(id);
  if (!state) {
    await answerCbq(token, cbq.id, 'Запрос устарел');
    return;
  }

  await answerCbq(token, cbq.id, '');
  pending.delete(id);

  try {
    if (type === 'po') {
      // Выбор объекта
      const objectName = state.objects[idx];
      if (!objectName) return reply('Ошибка выбора');
      await editKeyboard(token, chatId, messageId, `Объект: ${objectName}`);
      await proceedWithObject(objectName, state.orderNum, state.reason, chatId, threadId, token, cfg, state.supplierHint, state.dateHint);

    } else if (type === 'pi') {
      // Выбор накладной
      const invoice = state.invoices[idx];
      if (!invoice) return reply('Ошибка выбора');
      await editKeyboard(token, chatId, messageId, `Накладная: ${invoice.num}`);
      const ok = await markDeleted(invoice.num, state.reason, cfg);
      await reply(ok ? 'Принято' : 'Накладная не найдена');
    }
  } catch (err) {
    console.error('[aiHandler] handleCallbackQuery:', err.message);
    await reply('Повторите запрос');
  }
}

module.exports = { handleAiMessage, handleCallbackQuery };
