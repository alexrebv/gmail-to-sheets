const { getAuthClient, getSheetsClient } = require('./auth');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

async function ensureSheetExists(sheetName, headers) {
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID не задан в .env');

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === sheetName);

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] },
    });
    console.log(`Создан лист: ${sheetName}`);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [headers] },
    });
  }
}

/**
 * Дописывает строки в лист, пропуская дубли по колонке C (Номер заказа).
 * Перед записью читает все существующие номера и фильтрует новые.
 */
async function appendRowsToSheet(sheetName, rows) {
  if (!rows.length) return;
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID не задан в .env');

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  // Читаем всю колонку C — номера заказов уже в таблице
  const colC = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!C:C`,
  });

  const colValues = colC.data.values || [];

  // Множество уже существующих номеров (нормализованных)
  const norm = s => (s || '').toString().replace(/\s/g, '').toLowerCase().trim();
  const existingNums = new Set(
    colValues.slice(1).map(r => norm(r[0])).filter(Boolean)
  );

  // Последняя заполненная строка
  let lastDataRow = 1;
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i][0] && colValues[i][0].toString().trim() !== '') {
      lastDataRow = i + 1;
      break;
    }
  }

  // Фильтруем дубли — row[2] это номер заказа (колонка C, индекс 2)
  const newRows = rows.filter(row => {
    const num = norm(row[2]);
    if (!num) return true; // строки без номера пропускаем через
    if (existingNums.has(num)) {
      console.log(`  [skip] Дубль: ${row[2]}`);
      return false;
    }
    return true;
  });

  if (newRows.length === 0) {
    console.log(`Все ${rows.length} строк уже есть в "${sheetName}" — пропущено`);
    return;
  }

  const skipped = rows.length - newRows.length;
  if (skipped > 0) console.log(`Пропущено дублей: ${skipped}`);

  const startRow = lastDataRow + 1;

  // Форматируем даты
  const formattedRows = newRows.map(row =>
    row.map(cell => {
      if (cell instanceof Date) {
        return cell.toLocaleString('ru-RU', {
          day: '2-digit', month: '2-digit', year: 'numeric',
          hour: '2-digit', minute: '2-digit',
          timeZone: process.env.TIMEZONE || 'Europe/Moscow',
        });
      }
      return cell ?? '';
    })
  );

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A${startRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: formattedRows },
  });

  console.log(`Записано в "${sheetName}" начиная со строки ${startRow}: ${newRows.length} строк`);
}

module.exports = { ensureSheetExists, appendRowsToSheet };
