/**
 * pendingReport.js
 *
 * Генерирует Excel-отчёт «Не принятые накладные» и отправляет в Telegram.
 * Используется в /status_all и по расписанию CRON_STATUS.
 *
 * Колонки Excel: Поставщик | Дата заказа | Объект | Номер накладной | Статус
 */

const fs      = require('fs');
const path    = require('path');
const os      = require('os');
const https   = require('https');
const ExcelJS = require('exceljs');
const { getAuthClient, getSheetsClient } = require('./auth');
const { parseDateStr, todayDateKey }     = require('./dateUtils');

function cleanObjectName(name) {
  return (name || '').replace(/\s+(ФГ|ДР|DR|DP|GSW)\s*$/i, '').trim();
}

/**
 * Читает «Отправлен» и возвращает только непринятые строки прошлых дней.
 * Строки с K, начинающимся на '❌', и архивные — пропускаются.
 */
async function fetchPendingRows(cfg) {
  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  const SH_SENT        = cfg.SHEET_SENT     || 'Отправлен';
  const SH_ACC         = cfg.SHEET_ACCEPTED || 'Принят';

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const [sentRes, accRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_SENT}'!A2:L` }),
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_ACC}'!A2:D` }),
  ]);

  const sentRows = sentRes.data.values || [];
  const accRows  = accRes.data.values  || [];

  const norm = s => (s || '').toString().replace(/\s/g, '').toLowerCase().replace(/^#/, '').trim();

  const acceptedNums = new Set();
  accRows.forEach(r => {
    const num  = norm(r[2]);
    const type = (r[3] || '').toString().trim();
    if (num && type === 'Принят') acceptedNums.add(num);
  });

  const todayKey = todayDateKey();
  const pending  = [];

  for (const row of sentRows) {
    const obj       = cleanObjectName((row[1] || '').toString().trim()); // B
    const numRaw    = (row[2] || '').toString().trim();                  // C
    const num       = norm(numRaw);
    const supplier  = (row[4] || '').toString().trim();                  // E
    const dateRaw   = (row[5] || '').toString().trim();                  // F
    const archive   = (row[11] || '').toString().trim();                 // L
    const accStatus = (row[10] || '').toString().trim();                 // K

    if (!obj || !numRaw || !dateRaw) continue;
    if (archive.startsWith('Архив')) continue;
    if (accStatus.startsWith('❌')) continue;
    if (acceptedNums.has(num)) continue;

    const parsed = parseDateStr(dateRaw);
    if (!parsed) continue;
    const { dd, mm, yyyy } = parsed;
    const sortKey = `${yyyy}${mm}${dd}`;
    if (sortKey >= todayKey) continue; // только прошлые дни

    pending.push({
      supplier:  supplier || 'Без поставщика',
      dateStr:   `${dd}.${mm}.${yyyy.slice(-2)}`,
      dateSortKey: sortKey,
      object:    obj,
      orderNum:  numRaw,
      status:    'Не принято',
    });
  }

  // Сортировка: дата ↑, поставщик, объект
  pending.sort((a, b) =>
    a.dateSortKey.localeCompare(b.dateSortKey) ||
    a.supplier.localeCompare(b.supplier) ||
    a.object.localeCompare(b.object)
  );

  return pending;
}

/**
 * Создаёт xlsx-файл
 */
async function buildPendingExcel(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Не принятые накладные');

  ws.columns = [
    { header: 'Поставщик',        key: 'supplier',  width: 32 },
    { header: 'Дата заказа',      key: 'dateStr',   width: 14 },
    { header: 'Объект',           key: 'object',    width: 30 },
    { header: 'Номер накладной',  key: 'orderNum',  width: 22 },
    { header: 'Статус',           key: 'status',    width: 14 },
  ];

  // Заголовок
  ws.getRow(1).eachCell(cell => {
    cell.font      = { bold: true, name: 'Arial', size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.getRow(1).height = 18;

  // Строки
  for (const r of rows) {
    ws.addRow([r.supplier, r.dateStr, r.object, r.orderNum, r.status]);
  }

  // Итог
  ws.addRow([`Итого не принято: ${rows.length}`]);

  const now    = new Date().toISOString().slice(0, 10);
  const tmpPath = path.join(os.tmpdir(), `pending_${now}.xlsx`);
  await wb.xlsx.writeFile(tmpPath);
  return tmpPath;
}

/**
 * Отправляет xlsx-файл в Telegram
 */
async function sendDocument(token, chatId, threadId, filePath, caption) {
  const fileBuffer = fs.readFileSync(filePath);
  const fileName   = path.basename(filePath);
  const boundary   = '----TGBoundary' + Date.now();

  let body = '';
  body += `--${boundary}\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n${chatId}\r\n`;
  if (threadId) {
    const tid = parseInt(threadId);
    if (!isNaN(tid) && tid > 0) {
      body += `--${boundary}\r\nContent-Disposition: form-data; name="message_thread_id"\r\n\r\n${tid}\r\n`;
    }
  }
  if (caption) {
    body += `--${boundary}\r\nContent-Disposition: form-data; name="caption"\r\n\r\n${caption}\r\n`;
  }

  const bodyStart  = Buffer.from(body, 'utf-8');
  const fileHeader = Buffer.from(
    `--${boundary}\r\nContent-Disposition: form-data; name="document"; filename="${fileName}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`,
    'utf-8'
  );
  const bodyEnd   = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf-8');
  const totalBody = Buffer.concat([bodyStart, fileHeader, fileBuffer, bodyEnd]);

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.telegram.org',
      path:     `/bot${token}/sendDocument`,
      method:   'POST',
      headers: {
        'Content-Type':   `multipart/form-data; boundary=${boundary}`,
        'Content-Length': totalBody.length,
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        const r = JSON.parse(data);
        if (!r.ok) console.error('[pendingReport] Telegram error:', r.description);
        resolve(r);
      });
    });
    req.on('error', reject);
    req.write(totalBody);
    req.end();
  });
}

/**
 * Главная функция — генерирует и отправляет отчёт
 */
async function sendPendingReport(chatId, threadId, cfg) {
  const token = cfg.TELEGRAM_TOKEN;
  if (!token) return;

  const rows = await fetchPendingRows(cfg);

  if (rows.length === 0) {
    const { sendMessage } = require('./telegram');
    await sendMessage(token, chatId, '✅ Все накладные за прошлые дни приняты.', threadId, 0);
    return;
  }

  const filePath = await buildPendingExcel(rows);
  const now = new Date().toLocaleString('ru-RU', {
    timeZone: cfg.TIMEZONE || 'Europe/Moscow',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });

  await sendDocument(token, chatId, threadId, filePath,
    `Не принятые накладные - ${rows.length} шт.\nСформировано: ${now}`);

  try { fs.unlinkSync(filePath); } catch {}

  console.log(`[pendingReport] Отправлен файл: ${rows.length} строк → чат ${chatId}`);
}

module.exports = { sendPendingReport };
