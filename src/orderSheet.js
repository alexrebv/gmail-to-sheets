/**
 * orderSheet.js
 *
 * Записывает позиции заказов в лист «Позиции» Google Sheets.
 * Структура листа:
 *   A  Дата обработки
 *   B  Поставщик
 *   C  Объект
 *   D  Номер заказа
 *   E  Дата заказа
 *   F  Дата доставки
 *   G  Название товара
 *   H  Артикул
 *   I  Фасовка
 *   J  Кол-во
 *   K  Статус GO
 */

const { getAuthClient, getSheetsClient } = require('./auth');
const { ensureSheetExists } = require('./sheets');

const SHEET_NAME = 'Позиции';
const HEADERS = [
  'Дата обработки', 'Поставщик', 'Объект', 'Номер заказа',
  'Дата заказа', 'Дата доставки', 'Название товара',
  'Артикул', 'Фасовка', 'Кол-во', 'Статус GO',
];
const COL = {
  DATE: 0, SUPPLIER: 1, OBJECT: 2, ORDER_NUM: 3,
  ORDER_DATE: 4, DELIVERY: 5, NAME: 6,
  ARTICLE: 7, PACK: 8, QTY: 9, STATUS: 10,
};

/**
 * Записывает позиции всех заказов в лист «Позиции».
 * Пропускает уже существующие номера заказов (по колонке D).
 *
 * @param {Array} orders [{supplier, object, orderNumber, orderDate, deliveryDate, items}]
 * @param {object} cfg
 */
async function writeOrderItemsToSheet(orders, cfg) {
  if (!orders || orders.length === 0) return;

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  await ensureSheetExists(SHEET_NAME, HEADERS);

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  // Читаем существующие номера заказов чтобы не дублировать
  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!D2:D`,
  }).catch(() => ({ data: { values: [] } }));

  const existingNums = new Set(
    (existing.data.values || []).map(r => (r[0] || '').toString().trim())
  );

  const now = new Date();
  const rows = [];

  for (const order of orders) {
    if (existingNums.has(order.orderNumber)) {
      console.log(`[orderSheet] Пропуск дубликата: ${order.orderNumber}`);
      continue;
    }
    for (const item of order.items) {
      rows.push([
        now,
        order.supplier || '',
        order.object || '',
        order.orderNumber || '',
        order.orderDate || '',
        order.deliveryDate || '',
        item.desc || '',
        item.article || '',
        item.pack || '',
        item.qty || 0,
        'GO отправлен',
      ]);
    }
  }

  if (rows.length === 0) {
    console.log('[orderSheet] Нет новых позиций для записи');
    return;
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: rows },
  });

  console.log(`[orderSheet] Записано ${rows.length} позиций из ${orders.length} заказов`);
}

/**
 * Обновляет статус GO для заказа по номеру заказа.
 *
 * @param {string} orderNumber
 * @param {string} status
 * @param {object} cfg
 */
async function updateGoStatus(orderNumber, status, cfg) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!D2:K`,
  }).catch(() => ({ data: { values: [] } }));

  const rows = res.data.values || [];
  const updates = [];

  rows.forEach((row, i) => {
    if ((row[COL.ORDER_NUM - COL.DATE] || '').toString().trim() === orderNumber.toString().trim()) {
      updates.push({ range: `'${SHEET_NAME}'!K${i + 2}`, values: [[status]] });
    }
  });

  if (updates.length === 0) return;

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: updates },
  });

  console.log(`[orderSheet] Статус "${status}" для ${orderNumber} (${updates.length} строк)`);
}

module.exports = { writeOrderItemsToSheet, updateGoStatus };
