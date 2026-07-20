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

    const artM = rowHtml.match(/<td[^>]*class="column0 style12[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    if (!artM) continue;
    const article = cleanHtml(artM[1]);
    if (!article || isNaN(article)) continue;

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

const C = {
  GREY:  'FF989595',
  WHITE: 'FFFFFFFF',
  BLUE:  'FF6E85BF',
  LIGHT: 'FFEBEBEB',
  PINK:  'FFC9A2A2',
  BLACK: 'FF000000',
};

function border() {
  const s = { style: 'thin' };
  return { top: s, bottom: s, left: s, right: s };
}

function applyHeader(cell, bg, fgColor) {
  cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  cell.font   = { bold: true, color: { argb: fgColor || C.WHITE }, name: 'Calibri', size: 11 };
  cell.border = border();
  cell.alignment = { vertical: 'bottom' };
}

/**
 * Строит xlsx для одного поставщика со всеми его заказами.
 * @param {string} supplier
 * @param {Array}  orders — [{object, orderNumber, orderDate, deliveryDate, items, total}]
 * @returns {string} путь к временному файлу
 */
async function buildSupplierExcel(supplier, orders) {
  const wb = new ExcelJS.Workbook();
  const sheetName = supplier.replace(/[:\\\/\?\*\[\]]/g, ' ').substring(0, 31);
  const ws = wb.addWorksheet(sheetName);

  ws.columns = [
    { key: 'c0', width: 18 },
    { key: 'c1', width: 14 },
    { key: 'c2', width: 14 },
    { key: 'c3', width: 14 },
    { key: 'c4', width: 10 },
    { key: 'c5', width: 20 },
    { key: 'c6', width: 18 },
    { key: 'c7', width: 18 },
    { key: 'c8', width: 12 },
    { key: 'c9', width: 18 },
  ];

  for (const order of orders) {
    // ── Заголовок заказа ────────────────────────────────────────────────
    const r0 = ws.addRow([order.object, '', '', '', '', '', '', '', '', 'Заказ']);
    r0.getCell(1).font = { bold: true, size: 16, name: 'Calibri' };
    r0.getCell(10).font = { bold: true, size: 28, name: 'Calibri' };
    r0.height = 40;

    const r1 = ws.addRow(['', '', '', '', '', '', '', '', 'Дата', order.orderDate]);
    [9, 10].forEach(i => { r1.getCell(i).border = border(); });

    const r2 = ws.addRow(['', '', '', '', '', '', '', '', 'Заказ #', order.orderNumber]);
    [9, 10].forEach(i => { r2.getCell(i).border = border(); });

    ws.addRow([]);

    // ── Поставщик / Получатель ───────────────────────────────────────────
    const r4 = ws.addRow(['Поставщик', '', '', '', 'Получатель', '', '', '', '', '']);
    ws.mergeCells(r4.number, 1, r4.number, 3);
    ws.mergeCells(r4.number, 5, r4.number, 10);
    [1, 5].forEach(i => applyHeader(r4.getCell(i), C.GREY));

    const r5 = ws.addRow([supplier, '', '', '', order.object, '', '', '', '', '']);
    ws.mergeCells(r5.number, 1, r5.number, 3);
    ws.mergeCells(r5.number, 5, r5.number, 10);
    [1, 5].forEach(i => { r5.getCell(i).border = border(); });

    ws.addRow([]);

    // ── Дата доставки ────────────────────────────────────────────────────
    const r7 = ws.addRow(['Ожидаемая дата доставки', '', '', '', '', '', '', '', '', '']);
    ws.mergeCells(r7.number, 1, r7.number, 10);
    const r7c = r7.getCell(1);
    r7c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.GREY } };
    r7c.font   = { color: { argb: C.WHITE }, name: 'Calibri', size: 11 };
    r7c.border = border();
    r7c.alignment = { horizontal: 'center' };

    const r8 = ws.addRow([order.deliveryDate || '', '', '', '', '', '', '', '', '', '']);
    ws.mergeCells(r8.number, 1, r8.number, 10);
    const r8c = r8.getCell(1);
    r8c.border = border();
    r8c.alignment = { horizontal: 'center' };

    ws.addRow([]);

    // ── Заголовки колонок ────────────────────────────────────────────────
    const rh = ws.addRow([
      'Номер продукта #', 'Описание', '', '', 'Кол-во',
      'Упаковка', 'Цена вкл. НДС', 'Сумма вкл. НДС', 'НДС', 'Сумма без НДС',
    ]);
    ws.mergeCells(rh.number, 2, rh.number, 4);
    for (let i = 1; i <= 10; i++) {
      const c = rh.getCell(i);
      c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.GREY } };
      c.font   = { color: { argb: C.WHITE }, name: 'Calibri', size: 11 };
      c.border = border();
    }

    // ── Строки товаров ────────────────────────────────────────────────────
    for (const item of order.items) {
      const ri = ws.addRow([
        item.article, item.desc, '', '', item.qty,
        item.pack, item.price, item.sum, item.vat, item.net,
      ]);
      ws.mergeCells(ri.number, 2, ri.number, 4);
      for (let i = 1; i <= 10; i++) {
        const c = ri.getCell(i);
        c.border = border();
        if (i >= 7) {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.LIGHT } };
          c.numFmt = '#,##0.00';
        }
      }
    }

    // ── Итоги ─────────────────────────────────────────────────────────────
    const totals = [
      ['Промежуточный итог', order.total, '', ''],
      ['Стоимость перевозки', '-',         '', ''],
      ['Другое',             '-',         '', ''],
      ['Итог',               order.total, '', ''],
    ];
    for (let t = 0; t < totals.length; t++) {
      const [label, val] = totals[t];
      const rt = ws.addRow(['', '', '', '', '', '', label, val, '', '']);
      rt.getCell(7).border = border();
      const vc = rt.getCell(8);
      vc.border = border();
      if (t === 3) {
        // Итог — синий
        [7, 8, 9, 10].forEach(i => {
          const c = rt.getCell(i);
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.BLUE } };
          c.border = border();
        });
        vc.font = { bold: true, name: 'Calibri' };
        if (typeof val === 'number') vc.numFmt = '#,##0.00';
      } else {
        [8, 9, 10].forEach(i => {
          rt.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.LIGHT } };
          rt.getCell(i).border = border();
        });
      }
    }

    // Разделитель между заказами
    ws.addRow([]);
    ws.addRow([]);
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

/**
 * Главная точка входа — вызывается из gmail.js после парсинга писем.
 * @param {Array}  parsedOrders — [{supplier, object, orderNumber, orderDate, deliveryDate, items, total}]
 * @param {object} cfg
 */
async function sendOrderExcelReports(parsedOrders, cfg) {
  // Флаг быстрого отключения
  if ((cfg.ENABLE_ORDER_EXCEL || '').toString().toLowerCase() === 'false') {
    console.log('[orderExcel] Отключено (ENABLE_ORDER_EXCEL=false)');
    return;
  }

  const token    = cfg.TELEGRAM_TOKEN    || process.env.TELEGRAM_TOKEN;
  const chatId   = cfg.TELEGRAM_CHAT_ID  || process.env.TELEGRAM_CHAT_ID;
  const threadId = cfg.TELEGRAM_THREAD_ID || process.env.TELEGRAM_THREAD_ID || null;
  if (!token || !chatId) return;

  // Группируем по поставщику
  const bySupplier = new Map();
  for (const order of parsedOrders) {
    if (!order.supplier || order.items.length === 0) continue;
    if (!bySupplier.has(order.supplier)) bySupplier.set(order.supplier, []);
    bySupplier.get(order.supplier).push(order);
  }

  if (bySupplier.size === 0) {
    console.log('[orderExcel] Нет заказов с товарными позициями');
    return;
  }

  const now = new Date().toISOString().slice(0, 10);

  for (const [supplier, orders] of bySupplier) {
    try {
      const filePath = await buildSupplierExcel(supplier, orders);
      const totalItems = orders.reduce((s, o) => s + o.items.length, 0);
      const caption = `${supplier}\nЗаказов: ${orders.length} | Позиций: ${totalItems}\n${now}`;
      await sendDocument(token, chatId, threadId, filePath, caption);
      try { fs.unlinkSync(filePath); } catch {}
      console.log(`[orderExcel] Отправлен: ${supplier} (${orders.length} заказов, ${totalItems} позиций)`);
    } catch (e) {
      console.error(`[orderExcel] Ошибка для ${supplier}: ${e.message}`);
    }
  }
}

module.exports = { parseOrderItems, parseDeliveryDate, sendOrderExcelReports };
