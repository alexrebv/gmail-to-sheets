// PostgreSQL — постепенный переезд с Google Таблицы (в книге кончается лимит
// в 10 млн ячеек, а «Позиции»/«Story» растут каждый день).
//
// Этап 1: ДВОЙНАЯ ЗАПИСЬ. Таблица остаётся источником правды и читается как
// раньше, а сюда данные дублируются, чтобы сверить полноту без риска. Любая
// ошибка базы гасится и НЕ ломает основной поток (mirror()).
//
// Подключение: переменная DATABASE_URL. В Railway её надо связать со
// свойством базы (Variables → New Variable → ${{Postgres.DATABASE_URL}}).
// Если переменной нет, модуль полностью выключен и приложение работает
// по-старому.

let Pool = null;
try { ({ Pool } = require('pg')); } catch { /* пакет не установлен — режим выключен */ }

let _pool = null;
let _initDone = false;
let _initFailed = false;

function enabled() {
  return !!(Pool && process.env.DATABASE_URL);
}

function pool() {
  if (!enabled()) return null;
  if (!_pool) {
    _pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      // Railway отдаёт сертификат, который не проходит стандартную проверку.
      ssl: /localhost|127\.0\.0\.1/.test(process.env.DATABASE_URL) ? false : { rejectUnauthorized: false },
      max: 4,
      idleTimeoutMillis: 30000,
      connectionTimeoutMillis: 10000,
    });
    _pool.on('error', e => console.error('[db] pool error:', e.message));
  }
  return _pool;
}

async function query(sql, params) {
  const p = pool();
  if (!p) return { rows: [] };
  return p.query(sql, params);
}

// Схема. Всё через IF NOT EXISTS — безопасно вызывать на каждом старте.
// Ключи заказов/вычерков естественные (номер), поэтому повторная запись
// одной и той же строки не плодит дубли (ON CONFLICT DO NOTHING в мирроре).
const SCHEMA = `
CREATE TABLE IF NOT EXISTS vycherki (
  id            text PRIMARY KEY,
  date_str      text,
  time_str      text,
  supplier      text,
  object        text,
  invoice_date  text,
  invoice_number text,
  positions     jsonb,
  invoice_photos jsonb,
  vycherk_photos jsonb,
  video_id      text,
  status        text,
  reject_comment text,
  creator_login text,
  creator_name  text,
  edited_at     text,
  product_photos jsonb,
  marking_photos jsonb,
  extra_comment text,
  created_at    timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS vycherki_object_idx   ON vycherki (object);
CREATE INDEX IF NOT EXISTS vycherki_supplier_idx ON vycherki (supplier);
CREATE INDEX IF NOT EXISTS vycherki_status_idx   ON vycherki (status);

CREATE TABLE IF NOT EXISTS positions (
  order_number  text NOT NULL,
  supplier      text,
  object        text,
  order_date    text,
  delivery_date text,
  product_name  text,
  article       text,
  pack          text,
  qty           text,
  processed_at  text,
  created_at    timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (order_number, article, object)
);
CREATE INDEX IF NOT EXISTS positions_supplier_idx ON positions (supplier);
CREATE INDEX IF NOT EXISTS positions_object_idx   ON positions (object);
CREATE INDEX IF NOT EXISTS positions_delivery_idx ON positions (delivery_date);

CREATE TABLE IF NOT EXISTS hand_orders (
  ts            timestamptz NOT NULL DEFAULT now(),
  order_number  text NOT NULL,
  supplier      text,
  object        text,
  sup_name      text,
  sup_code      text,
  pack          text,
  qty           text,
  our_name      text,
  user_name     text,
  login         text,
  deadline      text
);
CREATE INDEX IF NOT EXISTS hand_orders_order_idx    ON hand_orders (order_number);
CREATE INDEX IF NOT EXISTS hand_orders_supplier_idx ON hand_orders (supplier);

CREATE TABLE IF NOT EXISTS entries (
  ts          timestamptz NOT NULL DEFAULT now(),
  tg_id       text,
  username    text,
  tg_name     text,
  login       text,
  role        text
);
CREATE INDEX IF NOT EXISTS entries_login_idx ON entries (login);
`;

async function init() {
  if (!enabled() || _initDone || _initFailed) return enabled() && _initDone;
  try {
    await query(SCHEMA);
    _initDone = true;
    console.log('[db] Postgres подключён, схема готова');
  } catch (e) {
    _initFailed = true;
    console.error('[db] не удалось подготовить схему:', e.message);
  }
  return _initDone;
}

// Зеркалирование: выполняет запись в базу best-effort. Никогда не бросает —
// пока Таблица остаётся источником правды, сбой базы не должен ничего ломать.
async function mirror(label, fn) {
  if (!enabled()) return;
  try {
    if (!_initDone && !_initFailed) await init();
    if (!_initDone) return;
    await fn();
  } catch (e) {
    console.error(`[db] mirror ${label}:`, e.message);
  }
}

const j = v => JSON.stringify(v ?? []);

// ── Зеркала конкретных сущностей ──────────────────────────────────────────────

async function mirrorVycherk(id, v, dateStr, timeStr) {
  await mirror('vycherk', () => query(
    `INSERT INTO vycherki (id, date_str, time_str, supplier, object, invoice_date, invoice_number,
       positions, invoice_photos, vycherk_photos, video_id, status, reject_comment,
       creator_login, creator_name, edited_at, product_photos, marking_photos, extra_comment)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19)
     ON CONFLICT (id) DO NOTHING`,
    [String(id), dateStr, timeStr, v.supplier || '', v.object || '', v.invoiceDate || '', v.invoiceNumber || '',
     j(v.positions), j(v.invoicePhotos), j(v.vycherkPhotos), v.videoId || '', 'Не обработан', '',
     v.creatorLogin || '', v.creatorName || '', '', j(v.productPhotos), j(v.markingPhotos), v.extraComment || '']
  ));
}

async function mirrorVycherkUpdate(id, fields) {
  const map = {
    status: 'status', rejectComment: 'reject_comment', positions: 'positions',
    invoicePhotos: 'invoice_photos', vycherkPhotos: 'vycherk_photos', videoId: 'video_id',
    supplier: 'supplier', object: 'object', invoiceDate: 'invoice_date', invoiceNumber: 'invoice_number',
    productPhotos: 'product_photos', markingPhotos: 'marking_photos', extraComment: 'extra_comment',
    editedAt: 'edited_at',
  };
  const sets = [], vals = [];
  for (const [k, col] of Object.entries(map)) {
    if (fields[k] === undefined) continue;
    vals.push(Array.isArray(fields[k]) ? j(fields[k]) : fields[k]);
    sets.push(`${col} = $${vals.length}`);
  }
  if (!sets.length) return;
  vals.push(String(id));
  await mirror('vycherk-update', () => query(`UPDATE vycherki SET ${sets.join(', ')} WHERE id = $${vals.length}`, vals));
}

async function mirrorHandOrder(rows) {
  if (!rows || !rows.length) return;
  await mirror('hand-order', async () => {
    for (const r of rows) {
      await query(
        `INSERT INTO hand_orders (ts, order_number, supplier, object, sup_name, sup_code, pack, qty, our_name, user_name, login, deadline)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
        [r[0] || new Date().toISOString(), r[10] || '', r[1] || '', r[2] || '', r[3] || '', r[4] || '',
         r[5] || '', r[6] || '', r[7] || '', r[8] || '', r[9] || '', r[11] || '']
      );
    }
  });
}

// Позиции заказов (лист «Позиции», основной источник роста Таблицы).
// rows — строки в порядке колонок листа:
// A дата обработки, B поставщик, C объект, D номер заказа, E дата заказа,
// F дата доставки, G товар, H артикул, I фасовка, J кол-во
async function mirrorPositions(rows) {
  if (!rows || !rows.length) return;
  await mirror('positions', async () => {
    for (const r of rows) {
      const orderNumber = String(r[3] ?? '').trim();
      const article = String(r[7] ?? '').trim();
      const object = String(r[2] ?? '').trim();
      if (!orderNumber || !object) continue;
      // DO UPDATE, а не DO NOTHING: заказ могут переобработать (reprocess),
      // и тогда количество/название должны обновиться, а не потеряться.
      await query(
        `INSERT INTO positions (order_number, supplier, object, order_date, delivery_date,
           product_name, article, pack, qty, processed_at)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)
         ON CONFLICT (order_number, article, object) DO UPDATE SET
           supplier = EXCLUDED.supplier, order_date = EXCLUDED.order_date,
           delivery_date = EXCLUDED.delivery_date, product_name = EXCLUDED.product_name,
           pack = EXCLUDED.pack, qty = EXCLUDED.qty, processed_at = EXCLUDED.processed_at`,
        [orderNumber, String(r[1] ?? '').trim(), object, String(r[4] ?? '').trim(), String(r[5] ?? '').trim(),
         String(r[6] ?? '').trim(), article, String(r[8] ?? '').trim(), String(r[9] ?? '').trim(),
         r[0] instanceof Date ? r[0].toISOString() : String(r[0] ?? '')]
      );
    }
  });
}

async function mirrorEntry(e) {
  await mirror('entry', () => query(
    `INSERT INTO entries (tg_id, username, tg_name, login, role) VALUES ($1,$2,$3,$4,$5)`,
    [String(e.tgId || ''), e.username || '', e.tgName || '', e.login || '', e.role || '']
  ));
}

// Сколько строк уже зеркалировано — для сверки с Таблицей.
async function counts() {
  if (!enabled()) return null;
  if (!_initDone && !_initFailed) await init();
  if (!_initDone) return null;
  const out = {};
  for (const t of ['vycherki', 'positions', 'hand_orders', 'entries']) {
    try { out[t] = Number((await query(`SELECT count(*)::int AS n FROM ${t}`)).rows[0].n); }
    catch { out[t] = null; }
  }
  return out;
}

module.exports = {
  enabled, init, query, mirror, counts,
  mirrorVycherk, mirrorVycherkUpdate, mirrorHandOrder, mirrorEntry, mirrorPositions,
};
