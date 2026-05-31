/**
 * channelBot.js
 *
 * Webhook-сервер для приёма сообщений из Telegram-канала @AcceptODChannel.
 * Парсит два типа сообщений и пишет в лист «Принят».
 *
 * ✅ ПРИНЯТ:
 * "Заказ принят!Заказ #202600-96-4654 ИП Власенко И.Ю. (поставка 30-05-2026)
 *  в ресторане DP Железнодорожная-02 был оприходован на склад"
 *
 * ❌ ОШИБКА (дата после слова "ошибкой"):
 * "Ошибка регистрации накладной!Регистрация накладной для заказа #20260-361-4794
 *  Скай ООО (RedBull) в ресторане OD Нахабино завершилась ошибкой (поставка 26-05-2026)"
 *
 * Колонки листа «Принят» (A–I):
 *   A  Дата записи       — когда пришло сообщение
 *   B  Поставщик         — ИП Власенко И.Ю.
 *   C  Номер заказа      — #202600-96-4654
 *   D  Тип               — Принят / Ошибка
 *   E  Дата поставки     — 30-05-2026
 *   F  Объект            — DP Железнодорожная-02 (без ФГ/ДР)
 *   G  Сырое сообщение   — полный текст для отладки
 *   H  (пусто)
 *   I  Статус проверки   — заполняет checkStatus.js
 */

const express = require('express');
const { getAuthClient, getSheetsClient } = require('./auth');
const { getConfig } = require('./config');
const { ensureSheetExists } = require('./sheets');

const app = express();
app.use(express.json());

// ── Regex-паттерны ────────────────────────────────────────────────────────────

// Принят: дата перед рестораном
const RE_ACCEPTED = /Заказ\s+(#\S+)\s+(.+?)\s+\(поставка\s+(\d{2}-\d{2}-\d{4})\)\s+в\s+ресторане\s+(.+?)\s+был\s+оприходован/i;

// Ошибка вариант A: дата ПОСЛЕ "ошибкой" — реальный формат канала
const RE_ERROR_A = /заказа\s+(#\S+)\s+(.+?)\s+в\s+ресторане\s+(.+?)\s+завершилась\s+ошибкой\s+\(поставка\s+(\d{2}-\d{2}-\d{4})\)/i;

// Ошибка вариант B: дата ДО ресторана — запасной вариант
const RE_ERROR_B = /заказа\s+(#\S+)\s+(.+?)\s+\(поставка\s+(\d{2}-\d{2}-\d{4})\)\s+в\s+ресторане\s+(.+?)\s+завершилась\s+ошибкой/i;

/** Убирает суффиксы ФГ, ДР и т.п. в конце названия объекта */
function cleanObjectName(name) {
  return (name || '').replace(/\s+(ФГ|ДР|DR|DP|GSW)\s*$/i, '').trim();
}

/**
 * Парсит текст сообщения из канала.
 * Возвращает { type, orderNumber, supplier, deliveryDate, object } или null.
 */
function parseChannelMessage(text) {
  if (!text) return null;

  // Нормализуем переносы строк
  const s = text.replace(/\n/g, ' ').replace(/\s{2,}/g, ' ');

  // --- Принят ---
  let m = s.match(RE_ACCEPTED);
  if (m) {
    return {
      type:         'Принят',
      orderNumber:  m[1].trim(),
      supplier:     m[2].trim(),
      deliveryDate: m[3].trim(),
      object:       cleanObjectName(m[4]),
    };
  }

  // --- Ошибка A: дата после "ошибкой" ---
  m = s.match(RE_ERROR_A);
  if (m) {
    return {
      type:         'Ошибка',
      orderNumber:  m[1].trim(),
      supplier:     m[2].trim(),
      object:       cleanObjectName(m[3]),
      deliveryDate: m[4].trim(),
    };
  }

  // --- Ошибка B: дата перед рестораном ---
  m = s.match(RE_ERROR_B);
  if (m) {
    return {
      type:         'Ошибка',
      orderNumber:  m[1].trim(),
      supplier:     m[2].trim(),
      deliveryDate: m[3].trim(),
      object:       cleanObjectName(m[4]),
    };
  }

  return null;
}

/**
 * Записывает одну строку в лист «Принят»
 */
async function writeToSheet(parsed, rawText, cfg) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SHEET_NAME     = cfg.SHEET_ACCEPTED || 'Принят';

  const HEADERS = [
    'Дата записи', 'Поставщик', 'Номер заказа', 'Тип',
    'Дата поставки', 'Объект', 'Сырое сообщение', '', 'Статус проверки',
  ];

  await ensureSheetExists(SHEET_NAME, HEADERS);

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  // Найти последнюю заполненную строку по колонке C (Номер заказа)
  const colC = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!C:C`,
  });
  const colValues = colC.data.values || [];
  let lastRow = 1;
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i][0] && colValues[i][0].toString().trim()) {
      lastRow = i + 1;
      break;
    }
  }

  const now = new Date().toLocaleString('ru-RU', {
    timeZone: cfg.TIMEZONE || 'Europe/Moscow',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  });

  const row = [
    now,                        // A — Дата записи
    parsed.supplier,            // B — Поставщик
    parsed.orderNumber,         // C — Номер заказа
    parsed.type,                // D — Тип
    parsed.deliveryDate,        // E — Дата поставки
    parsed.object,              // F — Объект
    rawText.substring(0, 500),  // G — Сырое сообщение
    '',                         // H — пусто
    '',                         // I — Статус проверки (заполнит checkStatus)
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A${lastRow + 1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [row] },
  });

  console.log(`[channelBot] ✓ ${parsed.type} | ${parsed.orderNumber} | ${parsed.object} | ${parsed.supplier}`);
}

// ── Webhook endpoint ──────────────────────────────────────────────────────────

app.post('/webhook', async (req, res) => {
  res.sendStatus(200); // отвечаем сразу — Telegram не будет повторять

  try {
    const update = req.body;
    const msg = update.channel_post || update.message;
    if (!msg) return;

    const text = msg.text || msg.caption || '';
    if (!text) return;

    const parsed = parseChannelMessage(text);
    if (!parsed) {
      console.log(`[channelBot] Пропущено: ${text.substring(0, 80)}`);
      return;
    }

    const cfg = await getConfig();
    await writeToSheet(parsed, text, cfg);

  } catch (err) {
    console.error(`[channelBot] Ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
});

app.get('/health', (req, res) =>
  res.json({ status: 'ok', ts: new Date().toISOString() }));

// ── Регистрация webhook в Telegram ───────────────────────────────────────────

async function registerWebhook(token, webhookUrl) {
  const https = require('https');
  const url = `https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(webhookUrl)}`;

  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        const r = JSON.parse(data);
        if (r.ok) console.log(`[channelBot] Webhook зарегистрирован: ${webhookUrl}`);
        else      console.error(`[channelBot] Ошибка webhook: ${r.description}`);
        resolve(r);
      });
    }).on('error', reject);
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
