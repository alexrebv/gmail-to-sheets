/**
 * pricelist.js
 *
 * Прайс-лист поставщика. Нужен там, где бланк заказа должен содержать весь
 * ассортимент, а не только то, что сегодня заказали: поставщик получает один
 * и тот же список в одном и том же порядке, а незаказанные позиции просто
 * остаются пустыми.
 *
 * Источников два, в этом порядке:
 *   1. База приложения (схема uchet, таблица price). Прайс заливают и правят
 *      в самом приложении, и это главный источник: с этими заказами потом
 *      работать, а в базе к каждой позиции лежит ещё и фасовка.
 *   2. Google Таблица - как было раньше. Остаётся запасным путём: база может
 *      быть не подключена к этому сервису, и бланк не должен из-за этого
 *      превращаться в обычный.
 *
 * Настройки (лист «Настройки», ключ → значение):
 *   FULL_BLANK_SUPPLIERS      — кому слать полный бланк, через запятую.
 *                               По умолчанию «Булки ПРО».
 *   SHEET_PRICELIST           — имя листа с прайсом (по умолчанию PriceList).
 *   PRICELIST_SPREADSHEET_ID  — если прайс лежит в другой таблице.
 *   PRICELIST_SOURCE          — «база» или «таблица», если нужно выбрать
 *                               источник руками. По умолчанию сначала база.
 *
 * Переменные окружения:
 *   UCHET_DATABASE_URL — база учёта, если она не та же, что DATABASE_URL.
 *
 * Колонки листа: A поставщик, B наш товар, C наш код, D товар у поставщика,
 * E код у поставщика, F фасовка. В бланк идут D, E и F - поставщик читает
 * свои названия, а не наши. В базе это те же поля: supplier, sup_name,
 * sup_code, pack.
 */

const { getAuthClient, getSheetsClient } = require('./auth');
const db = require('./db');

let Pool = null;
try { ({ Pool } = require('pg')); } catch { /* пакета нет - работаем по таблице */ }

const CACHE_TTL_MS = 5 * 60 * 1000;
let _cache = null;
let _cacheTime = 0;
let _usedSheet = null;   // какой лист прочитали на самом деле
let _usedSource = null;  // «база» или «таблица» - что сработало
let _dbNote = null;      // почему база не сработала, если не сработала
let _uchetPool = null;

// Латинские и русские двойники: «ООО» пишут и теми, и другими буквами, а
// глазом это не различить. Сводим к одному написанию, иначе поставщик из
// прайса не сойдётся с поставщиком из письма.
const LOOKALIKE = {
  a: 'а', c: 'с', e: 'е', o: 'о', p: 'р', x: 'х', y: 'у',
  b: 'в', h: 'н', k: 'к', m: 'м', t: 'т',
};

// Ключ поставщика: только буквы и цифры, в нижнем регистре, двойники сведены.
// «ООО «БУЛКИ  ПРО»» с двойным пробелом и «ООО Булки ПРО» - один поставщик.
function supKey(name) {
  return String(name || '').toLowerCase()
    .replace(/[a-z]/g, ch => LOOKALIKE[ch] || ch)
    .replace(/[^0-9a-zа-яё]+/g, '');
}

// Артикул приходит по-разному: «033» в письме и «33» в прайсе. Сравниваем без
// ведущих нулей и пробелов.
function artKey(code) {
  return String(code == null ? '' : code).trim().replace(/^0+/, '');
}

// Кому полный бланк. Пусто в настройках - «Булки ПРО»: с них всё началось,
// и молча перестать слать им полный список было бы хуже, чем настройка.
function fullBlankSuppliers(cfg) {
  const raw = (cfg && cfg.FULL_BLANK_SUPPLIERS) || 'Булки ПРО';
  return String(raw).split(',').map(s => s.trim()).filter(Boolean);
}

// В настройках пишут «Булки ПРО», а в письме приходит «ООО «БУЛКИ  ПРО»» -
// с формой собственности и кавычками. Поэтому сверяем вхождением, а не буква
// в букву; короткие обрывки («ооо») отбрасываем, чтобы не поймать всех подряд.
// Один и тот же поставщик, записанный по-разному. Сверяем вхождением:
// «Булки ПРО» в настройках и «ООО «БУЛКИ  ПРО»» в письме - одно и то же.
function sameSupplier(a, b) {
  const x = supKey(a), y = supKey(b);
  if (!x || !y) return false;
  if (x === y) return true;
  // Короткие обрывки («ооо») не считаем совпадением: так можно поймать всех.
  return Math.min(x.length, y.length) >= 4 && (x.includes(y) || y.includes(x));
}

function needsFullBlank(supplier, cfg) {
  if (!supKey(supplier)) return false;
  return fullBlankSuppliers(cfg).some(s => sameSupplier(supplier, s));
}

// Где искать прайс - одним местом, чтобы диагностика и чтение не разошлись.
function priceListSource(cfg) {
  return {
    spreadsheetId: (cfg && cfg.PRICELIST_SPREADSHEET_ID) || process.env.SPREADSHEET_ID || '',
    sheet: (cfg && cfg.SHEET_PRICELIST) || 'PriceList',
    ownId: !(cfg && cfg.PRICELIST_SPREADSHEET_ID),
  };
}

/** Названия листов книги - чтобы сказать, где искать, если заданного нет. */
async function sheetTitles(spreadsheetId) {
  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties.title',
  });
  return (res.data.sheets || []).map(x => x.properties && x.properties.title).filter(Boolean);
}

// Лист с прайсом, если имя не задано и «PriceList» в книге нет. Ищем по
// названию: «Прайс», «Прайс-лист», «PriceList», «Цены».
const PRICE_SHEET = /^\s*(прайс|price\s*list|pricelist|цены)/i;

// Какой источник просили. «база» / «db» - только база, «таблица» / «sheet» -
// только таблица, пусто - сначала база, потом таблица.
function wantSource(cfg) {
  const raw = String((cfg && cfg.PRICELIST_SOURCE) || '').trim().toLowerCase();
  if (/^(база|базы|db|database)/.test(raw)) return 'база';
  if (/^(таблиц|лист|sheet)/.test(raw)) return 'таблица';
  return 'авто';
}

// База учёта: прайс лежит в схеме uchet приложения. Обычно это та же база,
// что у нас (DATABASE_URL), но сервисы могут смотреть в разные - тогда путь
// задаётся отдельно, UCHET_DATABASE_URL.
function uchetUrl() {
  return process.env.UCHET_DATABASE_URL || process.env.DATABASE_URL || '';
}

function uchetPool() {
  const url = uchetUrl();
  if (!Pool || !url) return null;
  if (process.env.UCHET_DATABASE_URL && process.env.UCHET_DATABASE_URL !== process.env.DATABASE_URL) {
    if (!_uchetPool) {
      _uchetPool = new Pool({
        connectionString: url,
        ssl: /localhost|127\.0\.0\.1/.test(url) ? false : { rejectUnauthorized: false },
        max: 2, idleTimeoutMillis: 30000, connectionTimeoutMillis: 10000,
      });
      _uchetPool.on('error', e => console.error('[pricelist] pool error:', e.message));
    }
    return _uchetPool;
  }
  return null;   // та же база - читаем через общий db.query
}

/**
 * Прайс из базы приложения. Пусто (не подключена, схемы нет, прайс не залит) -
 * возвращаем null, чтобы вызвавший знал: надо смотреть таблицу.
 *
 * Порядок сразу алфавитный по названию у поставщика: в бланк позиции идут
 * этим же порядком, а в базе своего порядка строк нет.
 */
async function loadFromDb() {
  const sql = `SELECT coalesce(btrim(supplier), '')  AS supplier,
                      coalesce(btrim(sup_name), '')  AS sup_name,
                      coalesce(btrim(sup_code), '')  AS sup_code,
                      coalesce(btrim(pack), '')      AS pack,
                      coalesce(btrim(our_code), '')  AS our_code,
                      coalesce(btrim(our_name), '')  AS our_name,
                      coalesce(pack_factor, 1)       AS pack_factor,
                      coalesce(btrim(order_unit_raw), '')   AS order_unit,
                      coalesce(btrim(measure_unit_raw), '') AS measure_unit
                 FROM uchet.price
                ORDER BY 1, 2`;
  const p = uchetPool();
  let res;
  if (p) res = await p.query(sql);
  else if (db.enabled()) res = await db.query(sql);
  else { _dbNote = 'база не подключена (нет DATABASE_URL)'; return null; }

  const rows = [];
  for (const r of (res.rows || [])) {
    if (!r.supplier || !r.sup_name) continue;
    rows.push({
      supplier: r.supplier,
      article: r.sup_code,
      desc: r.sup_name,
      pack: r.pack,
      ourCode: r.our_code,
      ourName: r.our_name,
      packFactor: Number(r.pack_factor) || 1,
      orderUnit: r.order_unit,
      measureUnit: r.measure_unit,
    });
  }
  if (!rows.length) { _dbNote = 'в базе прайс пустой'; return null; }
  rows.sort((a, b) => a.supplier.localeCompare(b.supplier, 'ru') || a.desc.localeCompare(b.desc, 'ru'));
  return rows;
}

/** Весь прайс одним куском: [{ supplier, article, desc, pack }]. */
async function loadPriceList(cfg) {
  const now = Date.now();
  if (_cache && now - _cacheTime < CACHE_TTL_MS) return _cache;

  // Сначала база: прайс заливают в приложение, и это главный источник.
  const want = wantSource(cfg);
  _dbNote = null;
  if (want !== 'таблица') {
    try {
      const rows = await loadFromDb();
      if (rows) {
        _cache = rows; _cacheTime = now; _usedSource = 'база'; _usedSheet = null;
        return rows;
      }
    } catch (e) {
      _dbNote = e.message;
      console.error('[pricelist] прайс из базы не прочитался:', e.message);
    }
    if (want === 'база') {
      throw new Error('прайс из базы не прочитался: ' + (_dbNote || 'причина неизвестна'));
    }
  }

  // Дальше идёт таблица - отмечаем это до всех проверок: даже если таблица не
  // задана, источником была она, и подпись к бланку должна сказать именно так.
  _usedSource = 'таблица';
  const { spreadsheetId, sheet: wanted } = priceListSource(cfg);
  if (!spreadsheetId) return [];

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);

  // Заданного листа в книге может не быть - тогда Google отвечает «Unable to
  // parse range», и по этой строке не понять, что делать. Смотрим, какие листы
  // есть, и берём похожий на прайс; если и такого нет - говорим прямо.
  const titles = await sheetTitles(spreadsheetId);
  let sheet = titles.includes(wanted) ? wanted : titles.find(t => PRICE_SHEET.test(t));
  if (!sheet) {
    throw new Error(`листа «${wanted}» в книге нет. Листы: ${titles.join(', ') || 'книга пуста'}`);
  }
  _usedSheet = sheet;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${sheet}'!A2:F5000`,
  });

  const rows = [];
  for (const r of (res.data.values || [])) {
    const supplier = String((r && r[0]) || '').trim();
    const desc = String((r && r[3]) || '').trim();
    if (!supplier || !desc) continue;
    rows.push({
      supplier,
      article: String((r && r[4]) || '').trim(),
      desc,
      pack: String((r && r[5]) || '').trim(),
    });
  }
  _cache = rows;
  _cacheTime = now;
  return rows;
}

/**
 * Ассортимент поставщика в том порядке, в каком он лежит в прайсе - там он
 * уже расставлен по алфавиту, и переставлять его по-своему незачем.
 * Повторы по артикулу схлопываем: в бланке одна строка на позицию.
 */
async function catalogFor(supplier, cfg) {
  let all;
  try {
    all = await loadPriceList(cfg);
  } catch (e) {
    console.error(`[pricelist] Не удалось прочитать прайс: ${e.message}`);
    return [];
  }
  const seen = new Set();
  const out = [];
  for (const r of all) {
    if (!sameSupplier(r.supplier, supplier)) continue;
    const key = artKey(r.article) || r.desc.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ article: r.article, desc: r.desc, pack: r.pack });
  }
  return out;
}

/**
 * Что видно в прайсе: куда смотрим, сколько строк, какие поставщики.
 * Нужен для команды /pricelist и для подписи к бланку, когда прайс не сошёлся:
 * иначе полный бланк молча превращается в обычный, и причину не найти.
 */
async function diagnose(supplier, cfg) {
  const src = priceListSource(cfg);
  const out = { ...src, usedSheet: null, titles: [], error: null,
                rows: 0, suppliers: [], matched: [], mine: 0, wanted: fullBlankSuppliers(cfg),
                want: wantSource(cfg), source: null, dbNote: null,
                dbUrl: !!uchetUrl(), dbOwn: !!process.env.UCHET_DATABASE_URL };

  // Сначала читаем - по итогу и видно, что сработало. Названия листов нужны
  // только если дошло до таблицы, за ними лишний запрос в Google.
  // Названия листов книги - главная подсказка, когда заданного листа нет.
  // Спрашиваем их и при ошибке чтения: именно тогда они и нужны.
  const withTitles = async () => {
    if (_usedSource !== 'таблица' || !src.spreadsheetId) return;
    try { out.titles = await sheetTitles(src.spreadsheetId); } catch (e) { out.error = out.error || e.message; }
  };

  let all;
  try {
    all = await loadPriceList(cfg);
  } catch (e) {
    out.error = e.message;
    out.source = _usedSource;
    out.dbNote = _dbNote;
    await withTitles();
    return out;
  }
  out.source = _usedSource;
  out.dbNote = _dbNote;
  if (_usedSource === 'таблица' && !src.spreadsheetId) {
    out.error = 'Не задан ни PRICELIST_SPREADSHEET_ID, ни SPREADSHEET_ID';
    return out;
  }
  await withTitles();
  out.usedSheet = _usedSheet;
  out.rows = all.length;
  out.suppliers = [...new Set(all.map(r => r.supplier))].sort((a, b) => a.localeCompare(b, 'ru'));
  if (supplier) {
    // Как поставщик записан в самом прайсе: в настройках пишут кусок названия,
    // а в бланк идёт то, что стоит в прайсе.
    out.matched = out.suppliers.filter(x => sameSupplier(x, supplier));
    out.mine = (await catalogFor(supplier, cfg)).length;
  }
  return out;
}

// Откуда взяли прайс в последний раз: «база» или «таблица». null - ещё не
// читали. Нужно подписи к бланку: по числу позиций источник не угадать.
function lastSource() { return _usedSource; }

// Где прайс прочитали - одной строкой, для подписи и отчёта.
function sourceText(d) {
  if (d.source === 'база') return 'база приложения';
  const of = d.ownId ? 'основной таблицы' : 'таблицы ' + String(d.spreadsheetId).slice(0, 8) + '…';
  return `лист «${d.usedSheet || d.sheet}» ${of}`;
}

// Человеческим языком: что не так с прайсом. null - всё в порядке.
function whyEmpty(d, supplier) {
  const where = sourceText(d);
  if (d.error) return `прайс не прочитался (${where}): ${d.error}`;
  if (!d.rows) {
    return d.source === 'база'
      ? 'в базе нет строк прайса - загрузите прайс в приложении, раздел «Учёт»'
      : `в прайсе нет строк (${where}) - проверьте, что в колонке A поставщик, в D товар у поставщика`;
  }
  if (!d.mine) {
    const list = d.suppliers.slice(0, 8).join(', ');
    return `в прайсе нет позиций поставщика «${supplier}» (${where}). Есть: ${list}${d.suppliers.length > 8 ? ' и ещё ' + (d.suppliers.length - 8) : ''}`;
  }
  return null;
}

function clearCache() {
  _cache = null; _cacheTime = 0; _usedSheet = null; _usedSource = null; _dbNote = null;
}

module.exports = { catalogFor, loadPriceList, needsFullBlank, fullBlankSuppliers,
  priceListSource, diagnose, whyEmpty, sameSupplier, sheetTitles, supKey, artKey, clearCache,
  sourceText, wantSource, loadFromDb, lastSource };
