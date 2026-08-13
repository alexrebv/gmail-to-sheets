/**
 * backfill-db.js
 *
 * Разовая перезаливка истории листа «Позиции» в Postgres.
 * Двойная запись подхватывает только НОВЫЕ заказы, а в Таблице уже накоплена
 * история — её переносим этим скриптом.
 *
 * Запуск (в Railway: сервис gmail-to-sheets → Console):
 *     node src/backfill-db.js
 *
 * Безопасно запускать повторно: вставка идёт через ON CONFLICT DO UPDATE,
 * дубликатов не будет. Таблица не изменяется — только читается.
 */

const { getAuthClient, getSheetsClient } = require('./auth');
const db = require('./db');

const SHEET_NAME = 'Позиции';
const CHUNK = 500;

async function main() {
  if (!db.enabled()) {
    console.error('DATABASE_URL не задан — подключите базу и повторите.');
    process.exit(1);
  }
  if (!(await db.init())) {
    console.error('Не удалось подготовить схему.');
    process.exit(1);
  }

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  if (!SPREADSHEET_ID) { console.error('SPREADSHEET_ID не задан.'); process.exit(1); }

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);

  console.log(`Читаю лист «${SHEET_NAME}»...`);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A2:J`,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = res.data.values || [];
  console.log(`Строк в Таблице: ${rows.length}`);

  const before = (await db.counts()).positions;
  console.log(`В базе до заливки: ${before}`);

  let done = 0, skipped = 0;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const chunk = rows.slice(i, i + CHUNK).filter(r => {
      const okRow = String(r[3] ?? '').trim() && String(r[2] ?? '').trim();
      if (!okRow) skipped++;
      return okRow;
    });
    await db.mirrorPositions(chunk);
    done += chunk.length;
    const pct = Math.round(Math.min(i + CHUNK, rows.length) / rows.length * 100);
    console.log(`  ${Math.min(i + CHUNK, rows.length)}/${rows.length} (${pct}%)`);
  }

  const after = (await db.counts()).positions;
  console.log('');
  console.log(`Обработано строк : ${done}`);
  console.log(`Пропущено (без номера заказа или объекта): ${skipped}`);
  console.log(`В базе стало     : ${after} (было ${before})`);
  console.log('');
  console.log('Готово. Таблица не изменялась.');
  process.exit(0);
}

main().catch(e => { console.error('Ошибка:', e.message); process.exit(1); });
