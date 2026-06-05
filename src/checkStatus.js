/**
 * checkStatus.js  — аналог скриптов 2 и 3 (объединён)
 *
 * 1. Сверяет лист «Отправлен» с листами «Принят» и «Вычерк»
 * 2. Проставляет статусы K (статус приёмки) и J (вычерк) в «Отправлен»
 * 3. Проставляет «Проверено» в «Принят»!I и «Вычеркнуто» в «Вычерк»!I
 * 4. Архивирует строки через ARCHIVE_DELAY_DAYS дней
 * 5. Отправляет сводку в Telegram (один чат, одна подтема)
 *
 * Колонки «Отправлен» (1-based):
 *   B=2  Объект        F=6  Дата отправки   H=8  Направлено
 *   C=3  Номер заказа  G=7  Юр.лицо         J=10 Вычерк
 *   E=5  Поставщик     K=11 Статус приёмки  L=12 Архив
 *
 * Колонки «Принят» (1-based):
 *   C=3  Номер заказа   E=5  Дата заказа   F=6  Объект   I=9  Статус
 *
 * Колонки «Вычерк» (1-based):
 *   C=3  Номер заказа   I=9  Статус
 */

const { getAuthClient, getSheetsClient } = require('./auth');
const { getConfig } = require('./config');
const { parseDateStr, todayDateKey } = require('./dateUtils');
const { sendPendingReport } = require('./pendingReport');

const COL = {
  SENT: {
    OBJ: 1, NUM: 2, SUPPLIER: 4, DATE: 5, LEGAL: 6,
    SENT_STATUS: 7, CROSSED: 9, ACC_STATUS: 10, ARCHIVE: 11,
  },
  ACC:  { NUM: 2, TYPE: 3, DATE: 4, OBJ: 5, STATUS: 8 },
  CRS:  { NUM: 2, STATUS: 8 },
};

async function updateOrderStatusAndNotify() {
  const cfg = await getConfig();

  const SPREADSHEET_ID   = process.env.SPREADSHEET_ID;
  const SH_SENT          = cfg.SHEET_SENT     || 'Отправлен';
  const SH_ACC           = cfg.SHEET_ACCEPTED || 'Принят';
  const SH_CRS           = cfg.SHEET_CROSSED  || 'Вычерк';
  const CHAT_ID          = cfg.TELEGRAM_CHAT_ID;
  const THREAD_ID        = cfg.TELEGRAM_THREAD_ID || null;
  const ARCHIVE_DAYS     = Number(cfg.ARCHIVE_DELAY_DAYS || 2);

  if (!CHAT_ID) {
    console.error('[checkStatus] TELEGRAM_CHAT_ID не задан в Настройках');
    return;
  }

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);

  // Читаем все три листа
  const [sentRes, accRes, crsRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_SENT}'!A2:L` }),
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_ACC}'!A2:I` }),
    sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${SH_CRS}'!A2:I` }),
  ]);

  const sentRows = sentRes.data.values || [];
  const accRows  = accRes.data.values  || [];
  const crsRows  = crsRes.data.values  || [];

  const norm = s => (s || '').toString().replace(/\s/g, '').trim().replace(/^#/, '');
  const disp = s => (s || '').toString().trim();

  const todayKey = todayDateKey(); // YYYYMMDD in local time

  // Быстрый индекс по вычеркнутым
  const crossedSet = new Set();
  const crossedRowIdx = {}; // num → index в crsRows
  crsRows.forEach((r, i) => {
    const n = norm(r[COL.CRS.NUM]);
    if (n) { crossedSet.add(n); crossedRowIdx[n] = i; }
  });

  // Быстрый индекс принятых только по номеру заказа
  // Тип берём из колонки D (индекс 3): «Принят» или «Ошибка»
  const acceptedMap = {}; // num → index
  const errorMap    = {}; // num → index
  accRows.forEach((r, i) => {
    const n    = norm(r[COL.ACC.NUM]);           // C — Номер заказа
    const type = (r[COL.ACC.TYPE] || '').toString().trim(); // D — Тип
    if (!n) return;
    if (type === 'Принят') acceptedMap[n] = i;
    if (type === 'Ошибка')  errorMap[n]   = i;
  });

  // Накапливаем обновления для batchUpdate
  const sentUpdates = [];
  const accUpdates  = [];
  const crsUpdates  = [];


  for (let i = 0; i < sentRows.length; i++) {
    const row       = sentRows[i];
    const rowNum    = i + 2; // строка в таблице

    const obj       = disp(row[COL.SENT.OBJ]);
    const numRaw    = row[COL.SENT.NUM] || '';
    const num       = norm(numRaw);
    const supplier  = disp(row[COL.SENT.SUPPLIER]);
    const dateRaw   = row[COL.SENT.DATE];
    const legal     = disp(row[COL.SENT.LEGAL]);
    const archFlag  = disp(row[COL.SENT.ARCHIVE]);

    if (!obj || !numRaw || !dateRaw) continue;
    if (archFlag.startsWith('Архив')) continue;

    const accStatus = disp(row[COL.SENT.ACC_STATUS]);
    if (accStatus.startsWith('❌')) continue; // удалено вручную / ошибка IIKO

    const dateParts = parseDateStr(dateRaw);
    if (!dateParts) continue;
    const { dd: pdd, mm: pmm, yyyy: pyyyy } = dateParts;
    const sortKey = `${pyyyy}${pmm}${pdd}`;
    if (sortKey >= todayKey) continue; // только прошлые дни

    // daysDiff for archiving
    const dateSent = new Date(`${pyyyy}-${pmm}-${pdd}T00:00:00`);
    const today    = new Date(); today.setHours(0, 0, 0, 0);
    const daysDiff = Math.floor((today - dateSent) / 86400000);

    // --- Вычерк ---
    let hasCrossed = false;
    if (crossedSet.has(num)) {
      hasCrossed = true;
      sentUpdates.push({ range: `'${SH_SENT}'!J${rowNum}`, values: [['Вычерк']] });
      const ci = crossedRowIdx[num];
      crsUpdates.push({ range: `'${SH_CRS}'!I${ci + 2}`, values: [['Вычеркнуто']] });
    }

    // --- Принят ---
    // Сравниваем только по номеру заказа
    const accIdx = acceptedMap[num];
    const errIdx = errorMap[num];
    const foundAccepted = accIdx !== undefined;
    const foundError    = errIdx !== undefined;

    if (foundAccepted) {
      accUpdates.push({ range: `'${SH_ACC}'!I${accIdx + 2}`, values: [['Проверено']] });
      sentUpdates.push({ range: `'${SH_SENT}'!K${rowNum}`, values: [['✅ Оприходовано']] });
    } else if (foundError) {
      accUpdates.push({ range: `'${SH_ACC}'!I${errIdx + 2}`, values: [['Ошибка проверена']] });
      sentUpdates.push({ range: `'${SH_SENT}'!K${rowNum}`, values: [['❌ Ошибка регистрации']] });
    } else {
      if (!accStatus) {
        sentUpdates.push({ range: `'${SH_SENT}'!K${rowNum}`, values: [['⏳ Не оприходовано']] });
      }
    }

    // --- Архивация ---
    if (foundAccepted && daysDiff >= ARCHIVE_DAYS) {
      const archDate = formatDate(today);
      sentUpdates.push({ range: `'${SH_SENT}'!L${rowNum}`, values: [[`Архив ${archDate}`]] });
    }

  }

  // --- Применяем все обновления в таблицах ---
  const applyUpdates = async (data, sheetLabel) => {
    if (!data.length) return;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data },
    });
    console.log(`[checkStatus] Обновлено ${data.length} ячеек в "${sheetLabel}"`);
  };

  await applyUpdates(sentUpdates, SH_SENT);
  await applyUpdates(accUpdates,  SH_ACC);
  await applyUpdates(crsUpdates,  SH_CRS);

  // --- Telegram: Excel-отчёт с непринятыми ---
  await sendPendingReport(CHAT_ID, THREAD_ID, cfg);

  console.log('[checkStatus] Завершено.');
}

// ── helpers ───────────────────────────────────────────────────────────────────

function formatDate(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}.${mm}.${yy}`;
}

function escMd(s) {
  return (s || '').toString().replace(/[_*`[]/g, '\\$&');
}

module.exports = { updateOrderStatusAndNotify };
