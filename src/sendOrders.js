/**
 * sendOrders.js  — аналог скрипта 1
 *
 * Читает лист «Отправлен», группирует незаотправленные строки по дате+поставщику,
 * отправляет сводку в Telegram и проставляет статус «Направлено» в колонку H.
 *
 * Колонки листа «Отправлен» (1-based):
 *   A=1  Дата письма
 *   B=2  Объект
 *   C=3  Номер заказа
 *   D=4  Дата заказа
 *   E=5  Поставщик
 *   F=6  Дата отправки  ← dateRaw
 *   G=7  Юр.лицо
 *   H=8  Направлено     ← статус отправки в TG
 *   I=9  Тело письма
 *   J=10 Вычерк
 *   K=11 Статус приёмки
 *   L=12 Архив
 */

const { getAuthClient, getSheetsClient } = require('./auth');
const { getConfig } = require('./config');
const { sendLongMessage } = require('./telegram');

async function sendOrdersToTelegram() {
  const cfg = await getConfig();

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SHEET_NAME     = cfg.SHEET_SENT      || 'Отправлен';
  const TOKEN          = cfg.TELEGRAM_TOKEN;
  const CHAT_ID        = cfg.TELEGRAM_CHAT_ID;
  const THREAD_ID      = cfg.TELEGRAM_THREAD_ID || null;
  const MIN_DATE       = new Date(cfg.MIN_DATE  || '2025-09-16');
  const WAIT_MS        = Number(cfg.WAIT_MS     || 5000);
  const MAX_LEN        = Number(cfg.MAX_MESSAGE_LENGTH || 4000);

  if (!TOKEN || !CHAT_ID) {
    console.error('[sendOrders] TELEGRAM_TOKEN или TELEGRAM_CHAT_ID не заданы в Настройках');
    return;
  }

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A2:L`,
  });

  const rows = res.data.values || [];
  console.log(`[sendOrders] Строк для обработки: ${rows.length}`);

  // Группировка: { "DD.MM.YYYY|Поставщик": [{ object, rowIndex }] }
  const grouped = {};
  const statusUpdates = []; // { rowIndex, value }

  rows.forEach((row, idx) => {
    const rowIndex = idx + 2; // +2 из-за заголовка (строка 1) и смещения массива
    const object   = (row[1] || '').toString().trim();   // B
    const supplier = (row[4] || '').toString().trim();   // E
    const dateRaw  = row[5];                              // F
    const status   = (row[7] || '').toString().trim();   // H

    if (!object)  return;
    if (status)   return; // уже направлено

    const dateObj = dateRaw instanceof Date ? dateRaw : new Date(dateRaw);
    if (isNaN(dateObj)) return;
    if (dateObj < MIN_DATE) return;

    const dateFormatted = formatDate(dateObj);
    const key = `${dateFormatted}|${supplier}`;

    if (!grouped[key]) grouped[key] = [];
    grouped[key].push({ object, rowIndex });
  });

  const keys = Object.keys(grouped);
  if (keys.length === 0) {
    console.log('[sendOrders] Нет новых строк для отправки.');
    return;
  }

  for (const key of keys) {
    const [dateStr, supplier] = key.split('|');
    const items = grouped[key];

    let text = `*Заказ ${escMd(dateStr)}*\n*Поставщик:* ${escMd(supplier)}\n\n`;
    items.forEach(({ object }) => {
      text += `• ${escMd(object)}\n`;
    });

    await sendLongMessage(TOKEN, CHAT_ID, text, THREAD_ID, MAX_LEN, WAIT_MS);

    // Ставим статус «Направлено» в колонку H
    const updates = items.map(({ rowIndex }) => ({
      range: `'${SHEET_NAME}'!H${rowIndex}`,
      values: [['Направлено']],
    }));

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: updates,
      },
    });

    console.log(`[sendOrders] Отправлено ${items.length} заказов (${dateStr} / ${supplier})`);
  }

  console.log('[sendOrders] Завершено.');
}

// ── helpers ──────────────────────────────────────────────────────────────────

function formatDate(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

/** Экранирует спецсимволы Markdown v1 для Telegram */
function escMd(s) {
  return (s || '').toString().replace(/([_*[\]()~`>#+\-=|{}.!])/g, '\\$1');
}

module.exports = { sendOrdersToTelegram };
