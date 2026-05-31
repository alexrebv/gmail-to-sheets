const { getAuthClient, getSheetsClient } = require('./auth');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

/**
 * Убеждается, что лист с нужным именем существует.
 * Если нет — создаёт и добавляет заголовки.
 */
async function ensureSheetExists(sheetName, headers) {
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID не задан в .env');

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === sheetName);

  if (!exists) {
    // Создаём лист
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: sheetName } } }],
      },
    });
    console.log(`Создан лист: ${sheetName}`);

    // Записываем заголовки
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [headers] },
    });
  }
}

/**
 * Находит последнюю заполненную строку в колонке C (индекс 2),
 * затем дописывает новые строки строго после неё.
 */
async function appendRowsToSheet(sheetName, rows) {
  if (!rows.length) return;
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID не задан в .env');

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);

  // Читаем колонку C чтобы найти последнюю заполненную строку
  const colC = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!C:C`,
  });

  const colValues = colC.data.values || [];
  let lastDataRow = 1; // минимум — строка заголовка
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i][0] && colValues[i][0].toString().trim() !== '') {
      lastDataRow = i + 1;
      break;
    }
  }

  const startRow = lastDataRow + 1;

  // Преобразуем даты в строки для Sheets
  const formattedRows = rows.map(row =>
    row.map(cell => {
      if (cell instanceof Date) {
        // Формат: DD.MM.YYYY HH:MM
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

  console.log(`Записано в "${sheetName}" начиная со строки ${startRow}: ${rows.length} строк`);
}

module.exports = { ensureSheetExists, appendRowsToSheet };
