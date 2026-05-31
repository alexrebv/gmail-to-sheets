/**
 * errorsReport.js
 *
 * Генерирует Excel-файл с заказами у которых тип «Ошибка» в листе «Принят»
 * и отправляет его в Telegram как документ.
 *
 * Команда: /errors
 */

const fs   = require('fs');
const path = require('path');
const os   = require('os');
const https = require('https');
const ExcelJS = require('exceljs');
const { getAuthClient, getSheetsClient } = require('./auth');

/**
 * Читает ошибки из листа «Принят» (тип = «Ошибка»)
 * Возвращает массив строк для Excel
 */
async function getErrorRows(cfg) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_ACC = cfg.SHEET_ACCEPTED || 'Принят';

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SH_ACC}'!A2:I`,
  });

  const rows = res.data.values || [];

  return rows.filter(r => {
    const type = (r[3] || '').toString().trim(); // D — Тип
    return type === 'Ошибка';
  }).map(r => ({
    dateRecorded: r[0] || '',   // A — Дата записи
    supplier:     r[1] || '',   // B — Поставщик
    orderNumber:  r[2] || '',   // C — Номер заказа
    deliveryDate: r[4] || '',   // E — Дата поставки
    object:       r[5] || '',   // F — Объект
    rawMessage:   r[6] || '',   // G — Сырое сообщение
  }));
}

/**
 * Создаёт xlsx-файл с ошибками
 * Возвращает путь к временному файлу
 */
async function buildErrorsExcel(rows, cfg) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Ошибки регистрации');

  // Заголовки
  ws.columns = [
    { header: 'Дата записи',    key: 'dateRecorded', width: 20 },
    { header: 'Поставщик',      key: 'supplier',     width: 30 },
    { header: 'Номер заказа',   key: 'orderNumber',  width: 20 },
    { header: 'Дата поставки',  key: 'deliveryDate', width: 16 },
    { header: 'Объект',         key: 'object',       width: 30 },
    { header: 'Сообщение',      key: 'rawMessage',   width: 60 },
  ];

  // Стиль заголовка
  ws.getRow(1).eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0392B' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.getRow(1).height = 22;

  // Данные — группируем по дате поставки
  const byDate = {};
  rows.forEach(r => {
    const d = r.deliveryDate || 'Без даты';
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(r);
  });

  // Сортируем даты
  const sortedDates = Object.keys(byDate).sort((a, b) => {
    const pa = a.split('-').reverse().join('');
    const pb = b.split('-').reverse().join('');
    return pa.localeCompare(pb);
  });

  let rowNum = 2;
  for (const date of sortedDates) {
    // Строка-разделитель с датой
    const dateRow = ws.getRow(rowNum);
    dateRow.getCell(1).value = `📅 Поставка: ${date}`;
    dateRow.getCell(1).font = { bold: true, name: 'Arial', size: 10, color: { argb: 'FF7B2D00' } };
    dateRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
    ws.mergeCells(`A${rowNum}:F${rowNum}`);
    dateRow.height = 18;
    rowNum++;

    for (const r of byDate[date]) {
      const row = ws.getRow(rowNum);
      row.getCell(1).value = r.dateRecorded;
      row.getCell(2).value = r.supplier;
      row.getCell(3).value = r.orderNumber;
      row.getCell(4).value = r.deliveryDate;
      row.getCell(5).value = r.object;
      row.getCell(6).value = r.rawMessage;

      row.eachCell(cell => {
        cell.font = { name: 'Arial', size: 9 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF0F0' } };
        cell.alignment = { vertical: 'middle', wrapText: false };
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } },
          right:  { style: 'thin', color: { argb: 'FFDDDDDD' } },
        };
      });
      row.height = 18;
      rowNum++;
    }
  }

  // Итог
  ws.getRow(rowNum).getCell(1).value = `Всего ошибок: ${rows.length}`;
  ws.getRow(rowNum).getCell(1).font = { bold: true, name: 'Arial', size: 10 };

  // Сохраняем во временный файл
  const now = new Date().toISOString().slice(0, 10);
  const tmpPath = path.join(os.tmpdir(), `errors_${now}.xlsx`);
  await wb.xlsx.writeFile(tmpPath);
  return tmpPath;
}

/**
 * Отправляет файл в Telegram как документ
 */
async function sendDocument(token, chatId, threadId, filePath, caption) {
  const FormData = require('./formData'); // простая реализация ниже
  const fileBuffer = fs.readFileSync(filePath);
  const fileName   = path.basename(filePath);

  return new Promise((resolve, reject) => {
    const boundary = '----TGBoundary' + Date.now();

    let body = '';
    body += `--${boundary}\r\n`;
    body += `Content-Disposition: form-data; name="chat_id"\r\n\r\n${chatId}\r\n`;

    if (threadId) {
      const tid = parseInt(threadId);
      if (!isNaN(tid) && tid > 0) {
        body += `--${boundary}\r\n`;
        body += `Content-Disposition: form-data; name="message_thread_id"\r\n\r\n${tid}\r\n`;
      }
    }

    if (caption) {
      body += `--${boundary}\r\n`;
      body += `Content-Disposition: form-data; name="caption"\r\n\r\n${caption}\r\n`;
    }

    const bodyStart = Buffer.from(body, 'utf-8');
    const fileHeader = Buffer.from(
      `--${boundary}\r\nContent-Disposition: form-data; name="document"; filename="${fileName}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`,
      'utf-8'
    );
    const bodyEnd = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf-8');

    const totalBody = Buffer.concat([bodyStart, fileHeader, fileBuffer, bodyEnd]);

    const req = https.request({
      hostname: 'api.telegram.org',
      path: `/bot${token}/sendDocument`,
      method: 'POST',
      headers: {
        'Content-Type': `multipart/form-data; boundary=${boundary}`,
        'Content-Length': totalBody.length,
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        const r = JSON.parse(data);
        if (!r.ok) console.error('[errorsReport] Ошибка отправки файла:', r.description);
        resolve(r);
      });
    });
    req.on('error', reject);
    req.write(totalBody);
    req.end();
  });
}

/**
 * Основная функция — вызывается по команде /errors
 */
async function sendErrorsReport(chatId, threadId, cfg) {
  const token = cfg.TELEGRAM_TOKEN;

  const rows = await getErrorRows(cfg);

  if (rows.length === 0) {
    const { sendMessage } = require('./telegram');
    await sendMessage(token, chatId, '✅ Ошибок регистрации нет.', threadId, 0);
    return;
  }

  const filePath = await buildErrorsExcel(rows, cfg);
  const now = new Date().toLocaleString('ru-RU', {
    timeZone: cfg.TIMEZONE || 'Europe/Moscow',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });

  await sendDocument(
    token, chatId, threadId, filePath,
    `❌ Ошибки регистрации накладных — ${rows.length} шт.\nСформировано: ${now}`
  );

  // Удаляем временный файл
  try { fs.unlinkSync(filePath); } catch {}

  console.log(`[errorsReport] Отправлен файл с ${rows.length} ошибками в чат ${chatId}`);
}

module.exports = { sendErrorsReport };
