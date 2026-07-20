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

/**
 * Возвращает sheetId листа «Позиции».
 */
async function getSheetId(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });

  const sheet = meta.data.sheets.find(s => s.properties.title === SHEET_NAME);
  return sheet ? sheet.properties.sheetId : null;
}

/**
 * Удаляет строки с указанными номерами заказов из листа «Позиции».
 * Удаляет снизу вверх чтобы индексы не сдвигались.
 */
async function deleteOrderRows(sheets, spreadsheetId, orderNumbers) {
  if (!orderNumbers || orderNumbers.size === 0) return;

  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${SHEET_NAME}'!D2:D`,
  }).catch(() => ({ data: { values: [] } }));

  const rows = existing.data.values || [];
  // Собираем индексы строк (0-based от строки 2) которые надо удалить
  const toDelete = [];
  rows.forEach((r, i) => {
    if (orderNumbers.has((r[0] || '').toString().trim())) toDelete.push(i);
  });
  if (toDelete.length === 0) return;

  const sheetId = await getSheetId(sheets, spreadsheetId);
  if (sheetId === null) return;

  // Строим запросы удаления снизу вверх (чтобы индексы не сдвигались)
  const requests = toDelete
    .sort((a, b) => b - a)
    .map(i => ({
      deleteDimension: {
        range: { sheetId, dimension: 'ROWS', startIndex: i + 1, endIndex: i + 2 }, // +1 т.к. строка 1 — заголовок
      },
    }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests },
  });
  console.log(`[orderSheet] Удалено ${requests.length} старых строк`);
}

/**
 * Записывает позиции всех заказов в лист «Позиции».
 *
 * @param {Array}   orders  [{supplier, object, orderNumber, orderDate, deliveryDate, items}]
 * @param {object}  cfg
 * @param {boolean} overwrite — если true, удаляет существующие строки для этих заказов перед записью
 */
async function writeOrderItemsToSheet(orders, cfg, overwrite = false) {
  if (!orders || orders.length === 0) return;

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  await ensureSheetExists(SHEET_NAME, HEADERS);

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const orderNums = new Set(orders.map(o => (o.orderNumber || '').toString().trim()).filter(Boolean));

  if (overwrite) {
    // Удаляем существующие строки для этих заказов
    await deleteOrderRows(sheets, SPREADSHEET_ID, orderNums);
  } else {
    // Дедупликация: пропускаем заказы которые уже есть
    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!D2:D`,
    }).catch(() => ({ data: { values: [] } }));

    const existingNums = new Set(
      (existing.data.values || []).map(r => (r[0] || '').toString().trim())
    );
    orders = orders.filter(o => {
      if (existingNums.has((o.orderNumber || '').trim())) {
        console.log(`[orderSheet] Пропуск дубликата: ${o.orderNumber}`);
        return false;
      }
      return true;
    });
    if (orders.length === 0) {
      console.log('[orderSheet] Нет новых позиций для записи');
      return;
    }
  }

  const now = new Date();
  const rows = [];
  for (const order of orders) {
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
    console.log('[orderSheet] Нет позиций для записи');
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
 * Обновляет статус GO для заказа по номеру.
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
    if ((row[0] || '').toString().trim() === orderNumber.toString().trim()) {
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
