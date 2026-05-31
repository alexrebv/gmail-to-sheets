/**
 * config.js
 * Читает лист «Настройки» из Google Таблицы и возвращает объект с параметрами.
 * Кэширует результат на 5 минут чтобы не долбить API при каждом запуске.
 */

const { getAuthClient, getSheetsClient } = require('./auth');

let _cache = null;
let _cacheTime = 0;
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 минут

/**
 * Возвращает объект с настройками { КЛЮЧ: "значение" }
 */
async function getConfig() {
  const now = Date.now();
  if (_cache && now - _cacheTime < CACHE_TTL_MS) return _cache;

  const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID не задан в env');

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Настройки!A2:B50',
  });

  const rows = res.data.values || [];
  const cfg = {};
  for (const [key, value] of rows) {
    if (key && key.toString().trim()) {
      cfg[key.toString().trim()] = (value || '').toString().trim();
    }
  }

  _cache = cfg;
  _cacheTime = now;

  return cfg;
}

/** Сбросить кэш вручную */
function clearConfigCache() {
  _cache = null;
  _cacheTime = 0;
}

module.exports = { getConfig, clearConfigCache };
