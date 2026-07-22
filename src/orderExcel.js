/**
 * orderExcel.js
 *
 * Парсит позиции товаров из HTML письма iiko и формирует Excel-файл.
 * Один файл на поставщика — все его заказы из текущего batch складываются в один лист.
 *
 * Управление: cfg.ENABLE_ORDER_EXCEL = 'false' → полностью отключает модуль.
 */

const fs      = require('fs');
const path    = require('path');
const os      = require('os');
const https   = require('https');
const ExcelJS = require('exceljs');

// ── Парсинг HTML ──────────────────────────────────────────────────────────────

function cleanHtml(s) {
  return (s || '')
    .replace(/<[^>]+>/g, '')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .trim();
}

function parseNum(s) {
  const n = parseFloat((s || '').replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

/**
 * Извлекает дату доставки из HTML письма iiko.
 * Ищет строку row8: column0 style7 (центрированная ячейка с датой доставки).
 */
function parseDeliveryDate(html) {
  const m = html.match(/<td[^>]*class="column0 style7[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
  return m ? cleanHtml(m[1]) : '';
}

/**
 * Извлекает все товарные строки из HTML письма iiko.
 * Строки продуктов: column0 style12 (артикул — числовой), column1 style13 (описание colspan=3),
 * column4 style14 (кол-во), column5 style14 (упаковка), column6 style15 (цена),
 * column7 style16 (сумма вкл. НДС), column8 style16 (НДС), column9 style16 (без НДС).
 */
function parseOrderItems(html) {
  if (!html) return [];
  const items = [];
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch;

  while ((rowMatch = rowRegex.exec(html)) !== null) {
    const rowHtml = rowMatch[1];
    if (!rowHtml.includes('column0 style12')) continue;

    const artM = rowHtml.match(/<td[^>]*class="column0 style12[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    if (!artM) continue;
    const article = cleanHtml(artM[1]);
    if (!article) continue;

    const descM  = rowHtml.match(/<td[^>]*class="column1 style13[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    const qtyM   = rowHtml.match(/<td[^>]*class="column4 style14[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    const packM  = rowHtml.match(/<td[^>]*class="column5 style14[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    const priceM = rowHtml.match(/<td[^>]*class="column6 style15[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    const sumM   = rowHtml.match(/<td[^>]*class="column7 style16[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    const vatM   = rowHtml.match(/<td[^>]*class="column8 style16[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    const netM   = rowHtml.match(/<td[^>]*class="column9 style16[^"]*"[^>]*>([\s\S]*?)<\/td>/i);

    items.push({
      article,
      desc:  cleanHtml(descM?.[1]),
      qty:   parseNum(cleanHtml(qtyM?.[1])),
      pack:  cleanHtml(packM?.[1]),
      price: parseNum(cleanHtml(priceM?.[1])),
      sum:   parseNum(cleanHtml(sumM?.[1])),
      vat:   parseNum(cleanHtml(vatM?.[1])),
      net:   parseNum(cleanHtml(netM?.[1])),
    });
  }
  return items;
}

// ── Excel ─────────────────────────────────────────────────────────────────────

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
const TOTAL_FILL  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };

function thinBorder() {
  const s = { style: 'thin', color: { argb: 'FF000000' } };
  return { top: s, bottom: s, left: s, right: s };
}

function colLetter(n) {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

/**
 * Строит xlsx-пивот для одного поставщика.
 * Строки = продукты, столбцы = объекты, ячейки = кол-во.
 * @param {string} supplier
 * @param {Array}  orders — [{object, orderNumber, orderDate, deliveryDate, items, total}]
 * @returns {string} путь к временному файлу
 */
async function buildSupplierExcel(supplier, orders) {
  // ── Дата для имени листа — из первого заказа (дата доставки или заказа) ──
  const dateLabel = (orders[0].deliveryDate || orders[0].orderDate || '')
    .replace(/\//g, '.').replace(/:/g, '-').trim();

  const wb = new ExcelJS.Workbook();
  const sheetName = (dateLabel || new Date().toLocaleDateString('ru-RU')).substring(0, 31);
  const ws = wb.addWorksheet(sheetName);

  // ── Собираем все уникальные объекты (порядок из входных данных) ──────────
  const objectsOrdered = [];
  const objectSet = new Set();
  for (const o of orders) {
    if (o.object && !objectSet.has(o.object)) {
      objectSet.add(o.object);
      objectsOrdered.push(o.object);
    }
  }

  // ── Собираем все уникальные продукты (по артикулу) ───────────────────────
  const productMap = new Map(); // article → {desc, pack}
  for (const o of orders) {
    for (const it of o.items) {
      if (!productMap.has(it.article)) {
        productMap.set(it.article, { desc: it.desc, pack: it.pack });
      }
    }
  }

  // ── Пивот: article → object → qty ────────────────────────────────────────
  const pivot = new Map();
  for (const o of orders) {
    for (const it of o.items) {
      if (!pivot.has(it.article)) pivot.set(it.article, new Map());
      const cur = pivot.get(it.article).get(o.object) || 0;
      pivot.get(it.article).set(o.object, cur + it.qty);
    }
  }
  const articles = [...productMap.keys()];

  // ── Ширины столбцов ───────────────────────────────────────────────────────
  // A=Название, B=Артикул, C=Фасовка, D..=объекты, last=Итого
  const totalCols = 3 + objectsOrdered.length + 1;
  ws.getColumn(1).width = 36;
  ws.getColumn(2).width = 18;
  ws.getColumn(3).width = 32;
  for (let i = 4; i <= totalCols; i++) ws.getColumn(i).width = 14;

  // ── Строка 1: заголовок ───────────────────────────────────────────────────
  const headerValues = [
    `${supplier} - ${dateLabel}`,
    'Артикул',
    'Фасовка',
    ...objectsOrdered,
    'Итого',
  ];
  const rh = ws.addRow(headerValues);
  rh.height = 36;
  for (let i = 1; i <= totalCols; i++) {
    const c = rh.getCell(i);
    c.fill   = HEADER_FILL;
    c.font   = { bold: true, name: 'Calibri', size: 10 };
    c.border = thinBorder();
    c.alignment = { wrapText: true, vertical: 'middle', horizontal: i === 1 ? 'left' : 'center' };
  }

  // ── Строки продуктов ──────────────────────────────────────────────────────
  const dataStartRow = 2;
  for (let ai = 0; ai < articles.length; ai++) {
    const art  = articles[ai];
    const prod = productMap.get(art);
    const rowN = dataStartRow + ai;
    const objQtys = objectsOrdered.map(obj => pivot.get(art)?.get(obj) || null);
    const lastCol  = colLetter(totalCols);
    const firstObjCol = colLetter(4);

    const rowValues = [prod.desc, art, prod.pack, ...objQtys, { formula: `SUM(${firstObjCol}${rowN}:${colLetter(3 + objectsOrdered.length)}${rowN})` }];
    const ri = ws.addRow(rowValues);
    for (let i = 1; i <= totalCols; i++) {
      const c = ri.getCell(i);
      c.border = thinBorder();
      c.font   = { name: 'Calibri', size: 10 };
      if (i >= 4) {
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    }
  }

  // ── Строка итогов ────────────────────────────────────────────────────────
  const totalRowN = dataStartRow + articles.length;
  const totalValues = ['Итого', null, null];
  for (let ci = 4; ci <= totalCols; ci++) {
    const col = colLetter(ci);
    totalValues.push({ formula: `SUM(${col}${dataStartRow}:${col}${totalRowN - 1})` });
  }
  const rt = ws.addRow(totalValues);
  rt.height = 20;
  for (let i = 1; i <= totalCols; i++) {
    const c = rt.getCell(i);
    c.fill   = TOTAL_FILL;
    c.font   = { bold: true, name: 'Calibri', size: 10 };
    c.border = thinBorder();
    if (i >= 4) c.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  const now      = new Date().toISOString().slice(0, 10);
  const safeName = supplier.replace(/[^а-яёА-ЯЁa-zA-Z0-9]/g, '_').substring(0, 28);
  const tmpPath  = path.join(os.tmpdir(), `order_${safeName}_${now}.xlsx`);
  await wb.xlsx.writeFile(tmpPath);
  return tmpPath;
}

// ── Telegram sendDocument ─────────────────────────────────────────────────────

function sendDocument(token, chatId, threadId, filePath, caption) {
  const fileBuffer = fs.readFileSync(filePath);
  const fileName   = path.basename(filePath);
  const boundary   = '----OEBoundary' + Date.now();

  let body = '';
  body += `--${boundary}\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n${chatId}\r\n`;
  if (threadId) {
    const tid = parseInt(threadId);
    if (!isNaN(tid) && tid > 0)
      body += `--${boundary}\r\nContent-Disposition: form-data; name="message_thread_id"\r\n\r\n${tid}\r\n`;
  }
  if (caption)
    body += `--${boundary}\r\nContent-Disposition: form-data; name="caption"\r\n\r\n${caption}\r\n`;

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
      headers:  {
        'Content-Type':   `multipart/form-data; boundary=${boundary}`,
        'Content-Length': totalBody.length,
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); } catch { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(totalBody);
    req.end();
  });
}

function logTgResponse(supplier, resp) {
  if (!resp.ok) {
    console.error(`[orderExcel] Telegram error for ${supplier}: ${JSON.stringify(resp)}`);
  } else {
    console.log(`[orderExcel] Telegram ok for ${supplier}, message_id=${resp.result?.message_id}`);
  }
}

// ── Накопитель заказов за текущий день ────────────────────────────────────────
// supplier → { orders: [], date: 'YYYY-MM-DD' }
// При повторном поступлении новых объектов в тот же день — мержим и пересылаем.
const _dayAccumulator = new Map();

function _todayKey() {
  return new Date().toISOString().slice(0, 10);
}

/**
 * Добавляет новые заказы в накопитель.
 * Возвращает { allOrders, newObjects } для каждого поставщика.
 */
function _accumulateOrders(supplier, newOrders) {
  const today = _todayKey();
  if (!_dayAccumulator.has(supplier)) {
    _dayAccumulator.set(supplier, { orders: [], date: today });
  }
  const acc = _dayAccumulator.get(supplier);
  // Сброс на новый день
  if (acc.date !== today) {
    acc.orders = [];
    acc.date = today;
  }

  // Объекты которые уже были
  const existingObjects = new Set(acc.orders.map(o => o.object));
  // Новые объекты (ранее не приходившие от этого поставщика сегодня)
  const newObjects = newOrders
    .filter(o => !existingObjects.has(o.object))
    .map(o => o.object);

  // Добавляем только реально новые объекты (дедупликация по object)
  for (const o of newOrders) {
    if (!existingObjects.has(o.object)) {
      acc.orders.push(o);
      existingObjects.add(o.object);
    }
  }

  return { allOrders: acc.orders, newObjects };
}

/**
 * Главная точка входа — вызывается из gmail.js после парсинга писем.
 * @param {Array}  parsedOrders — [{supplier, object, orderNumber, orderDate, deliveryDate, items, total}]
 * @param {object} cfg
 */
async function sendOrderExcelReports(parsedOrders, cfg) {
  console.log(`[orderExcel] called: ${parsedOrders.length} orders, ENABLE_ORDER_EXCEL=${cfg.ENABLE_ORDER_EXCEL}`);

  // Флаг быстрого отключения
  if ((cfg.ENABLE_ORDER_EXCEL || '').toString().toLowerCase() === 'false') {
    console.log('[orderExcel] Отключено (ENABLE_ORDER_EXCEL=false)');
    return;
  }

  const token    = cfg.TELEGRAM_TOKEN    || process.env.TELEGRAM_TOKEN;
  const chatId   = cfg.TELEGRAM_CHAT_ID  || process.env.TELEGRAM_CHAT_ID;
  const threadId = cfg.TELEGRAM_THREAD_ID || process.env.TELEGRAM_THREAD_ID || null;
  console.log(`[orderExcel] token=${token ? 'set' : 'MISSING'} chatId=${chatId || 'MISSING'} threadId=${threadId}`);
  if (!token || !chatId) return;

  // Группируем по поставщику (если не определён — «Поставщик не определён»)
  const bySupplier = new Map();
  for (const order of parsedOrders) {
    if (order.items.length === 0) continue;
    if (!order.supplier) order.supplier = 'Поставщик не определён';
    if (!bySupplier.has(order.supplier)) bySupplier.set(order.supplier, []);
    bySupplier.get(order.supplier).push(order);
  }

  if (bySupplier.size === 0) {
    console.log('[orderExcel] Нет заказов с товарными позициями');
    return;
  }

  const now = _todayKey();

  for (const [supplier, newOrders] of bySupplier) {
    try {
      const isFirstBatch = !_dayAccumulator.has(supplier) ||
        _dayAccumulator.get(supplier).date !== now ||
        _dayAccumulator.get(supplier).orders.length === 0;

      const { allOrders, newObjects } = _accumulateOrders(supplier, newOrders);

      if (!isFirstBatch && newObjects.length === 0) {
        console.log(`[orderExcel] ${supplier}: нет новых объектов, пропускаем`);
        continue;
      }

      const filePath   = await buildSupplierExcel(supplier, allOrders);
      const totalItems = allOrders.reduce((s, o) => s + o.items.length, 0);

      let caption = `${supplier}\nЗаказов: ${allOrders.length} | Позиций: ${totalItems}\n${now}`;
      if (!isFirstBatch && newObjects.length > 0) {
        caption += `\n⚠️ Добавился объект: ${newObjects.join(', ')}`;
      }

      const resp = await sendDocument(token, chatId, threadId, filePath, caption);
      logTgResponse(supplier, resp);
      try { fs.unlinkSync(filePath); } catch {}
    } catch (e) {
      console.error(`[orderExcel] Ошибка для ${supplier}: ${e.message}`);
    }
  }
}

module.exports = { parseOrderItems, parseDeliveryDate, sendOrderExcelReports };
