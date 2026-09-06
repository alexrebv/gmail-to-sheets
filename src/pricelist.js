/**
 * pricelist.js
 *
 * Прайс-лист поставщика из Google Таблицы. Нужен там, где бланк заказа должен
 * содержать весь ассортимент, а не только то, что сегодня заказали: поставщик
 * получает один и тот же список в одном и том же порядке, а незаказанные
 * позиции просто остаются пустыми.
 *
 * Настройки (лист «Настройки», ключ → значение):
 *   FULL_BLANK_SUPPLIERS      — кому слать полный бланк, через запятую.
 *                               По умолчанию «Булки ПРО».
 *   SHEET_PRICELIST           — имя листа с прайсом (по умолчанию PriceList).
 *   PRICELIST_SPREADSHEET_ID  — если прайс лежит в другой таблице.
 *
 * Колонки листа: A поставщик, B наш товар, C наш код, D товар у поставщика,
 * E код у поставщика, F фасовка. В бланк идут D, E и F - поставщик читает
 * свои названия, а не наши.
 */

const { getAuthClient, getSheetsClient } = require('./auth');

const CACHE_TTL_MS = 5 * 60 * 1000;
let _cache = null;
let _cacheTime = 0;

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
function needsFullBlank(supplier, cfg) {
  const want = supKey(supplier);
  if (!want) return false;
  return fullBlankSuppliers(cfg).some(s => {
    const k = supKey(s);
    return k.length >= 4 && (want === k || want.includes(k) || k.includes(want));
  });
}

/** Весь прайс одним куском: [{ supplier, article, desc, pack }]. */
async function loadPriceList(cfg) {
  const now = Date.now();
  if (_cache && now - _cacheTime < CACHE_TTL_MS) return _cache;

  const spreadsheetId = (cfg && cfg.PRICELIST_SPREADSHEET_ID) || process.env.SPREADSHEET_ID;
  if (!spreadsheetId) return [];
  const sheet = (cfg && cfg.SHEET_PRICELIST) || 'PriceList';

  const auth = await getAuthClient();
  const sheets = getSheetsClient(auth);
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
  const want = supKey(supplier);
  const same = a => {
    const k = supKey(a);
    return k.length >= 4 && (k === want || k.includes(want) || want.includes(k));
  };
  const seen = new Set();
  const out = [];
  for (const r of all) {
    if (!same(r.supplier)) continue;
    const key = artKey(r.article) || r.desc.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ article: r.article, desc: r.desc, pack: r.pack });
  }
  return out;
}

function clearCache() { _cache = null; _cacheTime = 0; }

module.exports = { catalogFor, loadPriceList, needsFullBlank, fullBlankSuppliers, supKey, artKey, clearCache };
