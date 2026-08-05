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

// ── Бланк дня: устойчивое к перезапуску состояние ────────────────────────────
// Накопитель заказов за день в orderExcel.js живёт в памяти и теряется при
// рестарте/редеплое: после этого бланк уходил только с новым заказом и без
// пометки «Добавился объект». Поэтому состав бланка фиксируем в отдельном
// служебном листе (день + поставщик + номер заказа), а сами позиции при
// восстановлении дочитываем из «Позиции» по этим номерам.
//
// Лист «БланкиДня»: A День (YYYY-MM-DD) | B Поставщик | C Номер заказа

const BLANK_SHEET = 'БланкиДня';
const BLANK_HEADERS = ['День', 'Поставщик', 'Номер заказа'];

/**
 * Номера заказов, уже вошедшие в бланк этого поставщика за указанный день.
 * @returns {Promise<Set<string>>}
 */
async function loadBlankOrderNumbers(supplier, dayKey) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  await ensureSheetExists(BLANK_SHEET, BLANK_HEADERS);

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${BLANK_SHEET}'!A2:C`,
  }).catch(() => ({ data: { values: [] } }));

  const out = new Set();
  for (const r of (res.data.values || [])) {
    if ((r[0] || '').toString().trim() !== dayKey) continue;
    if ((r[1] || '').toString().trim() !== supplier) continue;
    const num = (r[2] || '').toString().trim();
    if (num) out.add(num);
  }
  return out;
}

/**
 * Дописывает номера заказов в состав бланка дня.
 */
async function recordBlankOrderNumbers(supplier, dayKey, orderNumbers) {
  const nums = [...new Set((orderNumbers || []).map(n => (n || '').toString().trim()).filter(Boolean))];
  if (!nums.length) return;

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  await ensureSheetExists(BLANK_SHEET, BLANK_HEADERS);

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${BLANK_SHEET}'!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: nums.map(n => [dayKey, supplier, n]) },
  });
}

/**
 * Восстанавливает заказы бланка из «Позиции» по номерам заказов.
 * Возвращает тот же формат, что и парсер писем:
 * [{supplier, object, orderNumber, orderDate, deliveryDate, items:[{desc,article,pack,qty}]}]
 */
async function loadOrdersByNumbers(orderNumbers) {
  const nums = orderNumbers instanceof Set ? orderNumbers : new Set(orderNumbers || []);
  if (nums.size === 0) return [];

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A2:J`,
  }).catch(() => ({ data: { values: [] } }));

  const byOrder = new Map();
  for (const r of (res.data.values || [])) {
    const num = (r[3] || '').toString().trim();
    if (!num || !nums.has(num)) continue;
    if (!byOrder.has(num)) {
      byOrder.set(num, {
        supplier:     (r[1] || '').toString().trim(),
        object:       (r[2] || '').toString().trim(),
        orderNumber:  num,
        orderDate:    (r[4] || '').toString().trim(),
        deliveryDate: (r[5] || '').toString().trim(),
        items: [],
      });
    }
    byOrder.get(num).items.push({
      desc:    (r[6] || '').toString().trim(),
      article: (r[7] || '').toString().trim(),
      pack:    (r[8] || '').toString().trim(),
      qty:     parseFloat((r[9] || '0').toString().replace(',', '.')) || 0,
    });
  }
  return [...byOrder.values()];
}

module.exports = {
  writeOrderItemsToSheet, updateGoStatus,
  loadBlankOrderNumbers, recordBlankOrderNumbers, loadOrdersByNumbers,
};
