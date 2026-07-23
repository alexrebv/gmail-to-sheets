const https = require('https');
const { getGmailClient, getAuthClient, getSheetsClient } = require('./auth');
const { appendRowsToSheet, ensureSheetExists } = require('./sheets');
const { getConfig } = require('./config');
const { parseOrderItems, parseDeliveryDate, sendOrderExcelReports } = require('./orderExcel');
const { writeOrderItemsToSheet } = require('./orderSheet');

/**
 * Убирает суффиксы ФГ, ДР, DR, DP, GSW в конце названия объекта
 * "OD Новохохловская-02 ФГ" → "OD Новохохловская-02"
 */
function cleanObjectName(name) {
  return (name || '').replace(/\s+(ФГ|ДР|DR|DP|GSW)\s*$/i, '').trim();
}

/**
 * Парсит тему письма
 * Поддерживает: #20260-749-0006, #2026-ОД15-3944
 * Время: 1:23 и 01:23
 */
function parseSubject(subject) {
  const regex = /Заказ для ресторана (.+?) (#\S+) создан (\d{2}\/\d{2}\/\d{2} \d{1,2}:\d{2})/;
  const match = subject.match(regex);
  if (match) {
    return {
      object:      cleanObjectName(match[1].trim()),
      orderNumber: match[2].trim(),
      orderDate:   match[3].trim(),
    };
  }
  return { object: '', orderNumber: '', orderDate: '' };
}

function decodeBody(data) {
  if (!data) return '';
  const base64 = data.replace(/-/g, '+').replace(/_/g, '/');
  return Buffer.from(base64, 'base64').toString('utf-8');
}

function decodeQuotedPrintable(str) {
  const clean = str.replace(/=\r?\n/g, '');
  const bytes = [];
  let i = 0;
  while (i < clean.length) {
    if (clean[i] === '=' && i + 2 < clean.length) {
      const hex = clean.slice(i + 1, i + 3);
      const byte = parseInt(hex, 16);
      if (!isNaN(byte)) {
        bytes.push(byte);
        i += 3;
        continue;
      }
    }
    const code = clean.charCodeAt(i);
    bytes.push(code > 127 ? 63 : code); // '?' for unexpected non-ASCII
    i++;
  }
  return Buffer.from(bytes).toString('utf-8');
}

function extractByMime(payload, mimeType) {
  if (!payload) return '';
  if (payload.mimeType === mimeType && payload.body?.data) {
    const raw = decodeBody(payload.body.data);
    const headers = payload.headers || [];
    // Определяем QP только по содержимому (soft line breaks =\n) —
    // заголовок CTE не используем: Gmail API может уже декодировать QP,
    // но оставлять заголовок quoted-printable, что приводит к двойному декоду.
    const isQP = /=\r?\n/.test(raw);
    const result = isQP ? decodeQuotedPrintable(raw) : raw;
    console.log(`[gmail] extractByMime ${mimeType}: len=${raw.length} isQP=${isQP}`);
    return result;
  }
  if (payload.parts) {
    for (const part of payload.parts) {
      const result = extractByMime(part, mimeType);
      if (result) return result;
    }
  }
  return '';
}

/**
 * Извлекает поставщика из HTML письма iiko.
 *
 * Структура письма:
 *   <tr> <td class="column0...">Поставщик</td> <td class="column4...">Получатель</td> </tr>
 *   <tr> <td class="column0...">ИП Григорян Рафик Айкович</td> ... </tr>
 *
 * Ищем ячейку column0 с текстом "Поставщик",
 * затем берём следующую строку column0 — это и есть имя поставщика.
 */
function extractSupplierFromHtml(html) {
  if (!html) return '';

  // Паттерн: блок "Поставщик" → следующая ячейка column0
  const pattern = /class="column0[^"]*"[^>]*>\s*Поставщик\s*<\/td>[\s\S]*?class="column0[^"]*"[^>]*>([\s\S]*?)<\/td>/i;
  const match = html.match(pattern);
  if (match) {
    const text = match[1]
      .replace(/<[^>]+>/g, '')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
      .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .trim();
    if (text && text.length > 2) return text;
  }

  // Fallback: ищем любую ячейку column0 с ИП/ООО/АО
  const cellRegex = /class="column0[^"]*"[^>]*>([\s\S]*?)<\/td>/gi;
  let m;
  while ((m = cellRegex.exec(html)) !== null) {
    const text = m[1]
      .replace(/<[^>]+>/g, '')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
      .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')
      .trim();
    if (/^(ИП|ООО|АО|ЗАО|ПАО)\s+/i.test(text)) return text;
  }

  return '';
}

/**
 * Извлекает итоговую сумму "Сумма вкл. НДС" из HTML письма iiko.
 *
 * Структура строки итогов в письме (column-индексы с нуля):
 *   column0..column5 — пустые ячейки
 *   column6 style4   — текст "Итог"
 *   column7 style18  — Сумма вкл. НДС  ← нужна эта
 *   column8 style18  — НДС
 *   column9 style18  — Сумма без НДС
 *
 * Стратегия: ищем <tr>, в котором есть ячейка с текстом "Итог" И ячейка
 * class="column7 style18". Берём значение из column7 style18.
 */
function extractOrderTotal(html) {
  if (!html) return null;

  // Ищем строку <tr> содержащую "Итог"
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(html)) !== null) {
    const rowHtml = rowMatch[1];

    // В строке должна быть ячейка с текстом "Итог"
    const rowText = rowHtml.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
    if (!/итог/i.test(rowText)) continue;

    // Ищем ячейку column7 style18 — это и есть "Сумма вкл. НДС"
    const col7Match = rowHtml.match(/<td[^>]*class="[^"]*column7[^"]*style18[^"]*"[^>]*>([\s\S]*?)<\/td>/i);
    if (col7Match) {
      const raw = col7Match[1].replace(/<[^>]+>/g, '').replace(/\s/g, '').replace(',', '.');
      const num = parseFloat(raw);
      if (!isNaN(num) && num > 0) return num;
    }

    // Fallback: берём первое положительное число в строке после ячейки "Итог"
    const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    let tdMatch;
    let pastItog = false;
    while ((tdMatch = tdRegex.exec(rowHtml)) !== null) {
      const cellText = tdMatch[1].replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
      if (/^итог$/i.test(cellText)) { pastItog = true; continue; }
      if (!pastItog) continue;
      const raw = cellText.replace(/\s/g, '').replace(',', '.');
      const num = parseFloat(raw);
      if (!isNaN(num) && num > 0) return num;
    }
  }

  return null;
}

// ── Кэш Распределение + Logins ───────────────────────────────────────────────

let _distCache = null, _distCacheTime = 0;
let _loginsCache = null, _loginsCacheTime = 0;
const DIST_TTL = 5 * 60 * 1000;

async function getDistribution(cfg) {
  if (_distCache && Date.now() - _distCacheTime < DIST_TTL) return _distCache;
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_DIST || 'Распределение'}'!A2:C`,
  });
  _distCache = (res.data.values || []).map(r => ({
    object: (r[0] || '').trim(),
    tu:     (r[1] || '').trim(),
    upr:    (r[2] || '').trim(),
  })).filter(d => d.object);
  _distCacheTime = Date.now();
  return _distCache;
}

async function getLogins(cfg) {
  if (_loginsCache && Date.now() - _loginsCacheTime < DIST_TTL) return _loginsCache;
  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_LOGINS || 'Logins'}'!A2:F`,
  });
  _loginsCache = (res.data.values || []).map(r => ({
    firstName: (r[0] || '').trim(),
    lastName:  (r[1] || '').trim(),
    tag:       (r[4] || '').trim().replace(/^@/, ''),
  })).filter(u => u.firstName);
  _loginsCacheTime = Date.now();
  return _loginsCache;
}

function findTagByName(fullName, loginsList) {
  const norm = s => s.toLowerCase().replace(/\s+/g, ' ').trim();
  const n    = norm(fullName);
  const user = loginsList.find(u => norm(`${u.firstName} ${u.lastName}`) === n);
  return user?.tag || '';
}

// ── Кэш поставщиков ───────────────────────────────────────────────────────────

let _suppliersCache    = null;
let _suppliersCacheTime = 0;
const SUPPLIERS_TTL    = 5 * 60 * 1000;

async function getSupplierMinimums(cfg) {
  if (_suppliersCache && Date.now() - _suppliersCacheTime < SUPPLIERS_TTL) return _suppliersCache;

  const auth   = await getAuthClient();
  const sheets = getSheetsClient(auth);
  const res    = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `'${cfg.SHEET_SUPPLIERS || 'Поставщики'}'!A2:B`,
  });

  const map = new Map();
  for (const row of (res.data.values || [])) {
    const name = (row[0] || '').toString().trim();
    const minVal = parseFloat((row[1] || '').toString().replace(/\s/g, '').replace(',', '.'));
    if (name && !isNaN(minVal)) map.set(name, minVal);
  }

  _suppliersCache    = map;
  _suppliersCacheTime = Date.now();
  return map;
}

function normSupplier(s) {
  return (s || '').toString().toLowerCase()
    .replace(/[«»"']/g, '').replace(/\s+/g, ' ').trim();
}

function findSupplierMin(supplierName, minimumsMap) {
  const norm = normSupplier(supplierName);
  for (const [key, val] of minimumsMap) {
    if (normSupplier(key) === norm) return val;
  }
  // Нечёткий поиск — вхождение
  for (const [key, val] of minimumsMap) {
    const k = normSupplier(key);
    if (norm.includes(k) || k.includes(norm)) return val;
  }
  return null;
}

// ── Telegram уведомление ──────────────────────────────────────────────────────

function tgPostGmail(token, chatId, threadId, text) {
  const body = JSON.stringify({
    chat_id: chatId,
    text,
    ...(threadId ? { message_thread_id: parseInt(threadId) } : {}),
  });
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.telegram.org',
      path:     `/bot${token}/sendMessage`,
      method:   'POST',
      headers:  { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => resolve(JSON.parse(d)));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

/** Fallback: поиск поставщика в plain-text теле */
function extractSupplierFromPlain(text) {
  if (!text) return '';
  for (const line of text.split('\n')) {
    const t = line.trim();
    if (/^(ИП|ООО|АО|ЗАО|ПАО)\s+/i.test(t)) {
      const dashIdx = t.indexOf(' - ');
      return dashIdx > 0 ? t.substring(0, dashIdx).trim() : t.trim();
    }
  }
  return '';
}

async function getOrCreateLabel(gmail, labelName) {
  const res = await gmail.users.labels.list({ userId: 'me' });
  const existing = (res.data.labels || []).find(l => l.name === labelName);
  if (existing) return existing.id;
  const created = await gmail.users.labels.create({
    userId: 'me',
    requestBody: { name: labelName },
  });
  console.log(`Создан лейбл: ${labelName}`);
  return created.data.id;
}

async function processGmailOrders() {
  try {
    const cfg = await getConfig();

    const LABEL_NAME        = cfg.GMAIL_LABEL  || 'Transfer';
    const SEARCH_QUERY_BASE = cfg.GMAIL_QUERY  || 'subject:"отправлен сотрудником"';
    const AFTER_DATE        = cfg.GMAIL_AFTER  || '2026/05/30';
    const SHEET_NAME        = cfg.SHEET_SENT   || 'Отправлен';

    const auth  = await getAuthClient();
    const gmail = getGmailClient(auth);

    const labelId   = await getOrCreateLabel(gmail, LABEL_NAME);
    const goLabelId = await getOrCreateLabel(gmail, 'GO');
    const query = `${SEARCH_QUERY_BASE} after:${AFTER_DATE} -label:${LABEL_NAME}`;
    console.log(`Поиск писем: ${query}`);

    const messageIds = [];
    let pageToken;
    do {
      const res = await gmail.users.messages.list({
        userId: 'me', q: query, maxResults: 100, pageToken,
      });
      const msgs = res.data.messages || [];
      messageIds.push(...msgs.map(m => m.id));
      pageToken = res.data.nextPageToken;
    } while (pageToken);

    console.log(`Найдено писем: ${messageIds.length}`);
    if (messageIds.length === 0) { console.log('Нет новых писем.'); return; }

    const newRows = [];
    const processedIds = [];
    const minimalkaRows = []; // строки для листа Минималка
    const orderExcelData = []; // данные для Excel-отчётов по поставщикам
    const orderEmailIds  = []; // id писем у которых распарсились позиции (для лейбла GO)

    const [supplierMins, dist, loginsList] = await Promise.all([
      getSupplierMinimums(cfg).catch(() => new Map()),
      getDistribution(cfg).catch(() => []),
      getLogins(cfg).catch(() => []),
    ]);

    for (const id of messageIds) {
      const res = await gmail.users.messages.get({ userId: 'me', id, format: 'full' });
      const msg     = res.data;
      const headers = msg.payload?.headers || [];

      const dateHeader = headers.find(h => h.name === 'Date')?.value || '';
      const subject    = headers.find(h => h.name === 'Subject')?.value || '';
      const emailDate  = dateHeader ? new Date(dateHeader) : new Date(parseInt(msg.internalDate));

      const htmlBody  = extractByMime(msg.payload, 'text/html');
      const plainBody = extractByMime(msg.payload, 'text/plain');

      console.log(`[gmail] htmlBody length: ${htmlBody.length}, hasQP: ${htmlBody.includes('=3D') || htmlBody.includes('=D0')}, sample: ${htmlBody.substring(0, 80).replace(/\n/g,' ')}`);

      const { object, orderNumber, orderDate } = parseSubject(subject);
      const supplier = extractSupplierFromHtml(htmlBody) || extractSupplierFromPlain(plainBody);
      const total    = extractOrderTotal(htmlBody);

      newRows.push([
        emailDate,         // A — Дата письма
        object || subject, // B — Объект (без ФГ)
        orderNumber,       // C — Номер заказа
        orderDate,         // D — Дата заказа
        supplier,          // E — Поставщик
        emailDate,         // F — Дата отправки
        '',                // G — Юр.лицо
        '',                // H — Направлено
        '',                // I — Тело (убрали HTML-мусор)
      ]);

      // Проверка минималки
      if (supplier && total !== null) {
        const minAmount = findSupplierMin(supplier, supplierMins);
        if (minAmount !== null) {
          const totalStr = total.toString().replace('.', ',');
          minimalkaRows.push({
            emailDate, object: object || subject, orderNumber, orderDate,
            supplier, total, totalStr, minAmount,
          });
        }
      }

      // Сбор позиций для Excel-отчёта
      const items        = parseOrderItems(htmlBody);
      const deliveryDate = parseDeliveryDate(htmlBody);
      console.log(`[gmail] parseOrderItems: ${items.length} items, deliveryDate: ${deliveryDate}`);
      if (items.length > 0) {
        orderExcelData.push({
          supplier: supplier || '',
          object:   object || subject,
          orderNumber, orderDate,
          deliveryDate, items, total,
        });
        orderEmailIds.push(id);
      }

      processedIds.push(id);
      console.log(`  ✓ ${object} | ${orderNumber} | ${supplier}${total !== null ? ` | сумма: ${total}` : ''}${items.length > 0 ? ` | позиций: ${items.length}` : ''}`);
    }

    const HEADERS = [
      'Дата письма', 'Объект', 'Номер заказа', 'Дата заказа',
      'Поставщик', 'Дата отправки', 'Юр.лицо', 'Направлено', 'Тело письма',
    ];
    await ensureSheetExists(SHEET_NAME, HEADERS);
    await appendRowsToSheet(SHEET_NAME, newRows);

    // ── Обработка Минималки ──────────────────────────────────────────────────
    if (minimalkaRows.length > 0) {
      const MSHEET  = cfg.SHEET_MINIMALKA || 'Минималка';
      const MHEADERS = ['Дата письма', 'Объект', 'Номер заказа', 'Дата заказа', 'Поставщик', 'Дата отправки', 'Сумма итог', 'Статус'];
      await ensureSheetExists(MSHEET, MHEADERS);

      const mRows = minimalkaRows.map(r => [
        r.emailDate, r.object, r.orderNumber, r.orderDate,
        r.supplier, r.emailDate, r.totalStr, 'обработан',
      ]);
      await appendRowsToSheet(MSHEET, mRows);

      // Уведомления по накладным ниже минималки
      const token    = cfg.TELEGRAM_TOKEN;
      const chatId   = cfg.TELEGRAM_CHAT_ID;
      const threadId = cfg.TELEGRAM_THREAD_ID || null;

      if (token && chatId) {
        for (const r of minimalkaRows) {
          if (r.total < r.minAmount) {
            const minStr = r.minAmount.toString().replace('.', ',');

            // Ищем управляющего (Upr, столбец C) или ТУ (столбец B) для объекта
            const distRow = dist.find(d => d.object === r.object);
            let tagStr = '';
            if (distRow) {
              const uprTag = distRow.upr ? findTagByName(distRow.upr, loginsList) : '';
              const tuTag  = distRow.tu  ? findTagByName(distRow.tu,  loginsList) : '';
              const tag    = uprTag || tuTag;
              if (tag) tagStr = `@${tag} `;
            }

            const text = `${tagStr}${r.object}\n${r.supplier}\nСумма: ${r.totalStr}\n\nМинималка не набрана\nНужно ${minStr}`;
            await tgPostGmail(token, chatId, threadId, text).catch(e =>
              console.error(`[gmail] Telegram notify error: ${e.message}`)
            );
            console.log(`[gmail] Минималка не набрана: ${r.object} | ${r.supplier} | ${r.total} < ${r.minAmount}`);
          }
        }
      }
    }
    console.log(`Записано строк: ${newRows.length}`);

    // ── Excel-отчёты по поставщикам + запись в лист Позиции ─────────────────
    // writeOrderItemsToSheet сам дедуплицирует по номеру заказа (overwrite=false
    // по умолчанию), так что дублей при параллельных запусках с reprocessTodayOrders
    // не будет — та функция явно исключает уже помеченные GO письма (-label:GO).
    if (orderExcelData.length > 0) {
      await sendOrderExcelReports(orderExcelData, cfg).catch(e =>
        console.error(`[gmail] orderExcel error: ${e.message}`)
      );
      await writeOrderItemsToSheet(orderExcelData, cfg).catch(e =>
        console.error(`[gmail] orderSheet error: ${e.message}`)
      );
    }

    for (const id of processedIds) {
      const addLabelIds = [labelId];
      if (orderEmailIds.includes(id)) addLabelIds.push(goLabelId);
      await gmail.users.messages.modify({
        userId: 'me', id,
        requestBody: { addLabelIds },
      });
    }
    console.log(`Помечено писем лейблом "${LABEL_NAME}": ${processedIds.length}, лейблом "GO": ${orderEmailIds.length}`);

  } catch (err) {
    console.error(`[processGmailOrders] Ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
}

/**
 * Ищет сегодняшние письма с лейблом Transfer но без GO.
 * Формирует Excel и пишет в лист «Позиции» для тех, у кого распарсились позиции.
 */
async function reprocessTodayOrders() {
  try {
    const cfg = await getConfig();

    const LABEL_NAME = cfg.GMAIL_LABEL || 'Transfer';
    const auth  = await getAuthClient();
    const gmail = getGmailClient(auth);

    const goLabelId = await getOrCreateLabel(gmail, 'GO');

    // Сегодняшняя дата для Gmail-запроса (YYYY/MM/DD)
    const today = new Date();
    const yyyy  = today.getFullYear();
    const mm    = String(today.getMonth() + 1).padStart(2, '0');
    const dd    = String(today.getDate()).padStart(2, '0');
    const todayStr = `${yyyy}/${mm}/${dd}`;

    const query = `label:${LABEL_NAME} -label:GO after:${todayStr}`;
    console.log(`[reprocess] Поиск сегодняшних без GO: ${query}`);

    const messageIds = [];
    let pageToken;
    do {
      const res = await gmail.users.messages.list({
        userId: 'me', q: query, maxResults: 100, pageToken,
      });
      messageIds.push(...(res.data.messages || []).map(m => m.id));
      pageToken = res.data.nextPageToken;
    } while (pageToken);

    console.log(`[reprocess] Найдено без GO: ${messageIds.length}`);
    if (messageIds.length === 0) return;


    const orderExcelData = [];
    const goIds = [];

    for (const id of messageIds) {
      const res  = await gmail.users.messages.get({ userId: 'me', id, format: 'full' });
      const msg  = res.data;
      const headers = msg.payload?.headers || [];
      const subject = headers.find(h => h.name === 'Subject')?.value || '';

      const htmlBody = extractByMime(msg.payload, 'text/html');
      const plainBody = extractByMime(msg.payload, 'text/plain');
      const { object, orderNumber, orderDate } = parseSubject(subject);
      const supplier = extractSupplierFromHtml(htmlBody) || extractSupplierFromPlain(plainBody);
      const total = extractOrderTotal(htmlBody);

      const items = parseOrderItems(htmlBody);
      const deliveryDate = parseDeliveryDate(htmlBody);

      const hasStyle12 = htmlBody.includes('column0 style12');
      console.log(`[reprocess] ${object} | ${orderNumber} | ${supplier} | items: ${items.length} | hasStyle12: ${hasStyle12} | htmlLen: ${htmlBody.length}`);
      if (items.length === 0 && htmlBody.length > 100) {
        // Показываем какие column-классы есть в HTML для диагностики
        const foundClasses = [...new Set((htmlBody.match(/class="column\d+ style\d+[^"]*"/g) || []))].slice(0, 10);
        console.log(`[reprocess] classes in html: ${foundClasses.join(', ')}`);
      }

      if (items.length > 0) {
        orderExcelData.push({ supplier, object, orderNumber, orderDate, deliveryDate, items, total });
        goIds.push(id);
      }
    }

    if (orderExcelData.length === 0) {
      console.log('[reprocess] Позиций не найдено');
      return;
    }

    await sendOrderExcelReports(orderExcelData, cfg).catch(e =>
      console.error(`[reprocess] orderExcel error: ${e.message}`)
    );
    // overwrite=true: удаляем старые строки этих заказов перед записью (чистый перезапуск)
    await writeOrderItemsToSheet(orderExcelData, cfg, true).catch(e =>
      console.error(`[reprocess] orderSheet error: ${e.message}`)
    );

    // Ставим лейбл GO на обработанные письма
    for (const id of goIds) {
      await gmail.users.messages.modify({
        userId: 'me', id,
        requestBody: { addLabelIds: [goLabelId] },
      }).catch(e => console.error(`[reprocess] label GO error: ${e.message}`));
    }
    console.log(`[reprocess] GO поставлен на ${goIds.length} писем`);

  } catch (err) {
    console.error(`[reprocessTodayOrders] Ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
}

/**
 * Разовый бэкфилл: перечитывает уже помеченные Transfer+GO письма за период
 * (по умолчанию с 2026/07/20 — дата коммита e498a40, который отключил запись
 * в лист «Позиции» для писем, распознавшихся с первого раза) и дописывает
 * недостающие позиции. Excel-отчёты и лейблы не трогает — эти письма уже
 * обработаны и разосланы в своё время, тут чинится только пробел в таблице.
 * Безопасно перезапускать: writeOrderItemsToSheet(..., overwrite=false)
 * пропускает номера заказов, которые уже есть в листе.
 * Включается переменной окружения BACKFILL_POSITIONS=true (см. index.js).
 */
async function backfillPositions() {
  try {
    const cfg = await getConfig();
    const LABEL_NAME = cfg.GMAIL_LABEL || 'Transfer';
    const AFTER_DATE = process.env.BACKFILL_AFTER_DATE || '2026/07/20';

    const auth  = await getAuthClient();
    const gmail = getGmailClient(auth);

    const query = `label:${LABEL_NAME} label:GO after:${AFTER_DATE}`;
    console.log(`[backfill] Поиск: ${query}`);

    const messageIds = [];
    let pageToken;
    do {
      const res = await gmail.users.messages.list({
        userId: 'me', q: query, maxResults: 100, pageToken,
      });
      messageIds.push(...(res.data.messages || []).map(m => m.id));
      pageToken = res.data.nextPageToken;
    } while (pageToken);

    console.log(`[backfill] Найдено писем с лейблом GO: ${messageIds.length}`);
    if (messageIds.length === 0) return;

    const orderExcelData = [];

    for (const id of messageIds) {
      const res     = await gmail.users.messages.get({ userId: 'me', id, format: 'full' });
      const msg     = res.data;
      const headers = msg.payload?.headers || [];
      const subject = headers.find(h => h.name === 'Subject')?.value || '';

      const htmlBody  = extractByMime(msg.payload, 'text/html');
      const plainBody = extractByMime(msg.payload, 'text/plain');
      const { object, orderNumber, orderDate } = parseSubject(subject);
      const supplier = extractSupplierFromHtml(htmlBody) || extractSupplierFromPlain(plainBody);
      const total    = extractOrderTotal(htmlBody);

      const items        = parseOrderItems(htmlBody);
      const deliveryDate = parseDeliveryDate(htmlBody);

      if (items.length > 0) {
        orderExcelData.push({
          supplier: supplier || '',
          object:   object || subject,
          orderNumber, orderDate, deliveryDate, items, total,
        });
      }
    }

    console.log(`[backfill] Писем с распознанными позициями: ${orderExcelData.length} из ${messageIds.length}`);
    if (orderExcelData.length === 0) return;

    await writeOrderItemsToSheet(orderExcelData, cfg);

  } catch (err) {
    console.error(`[backfillPositions] Ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
}

module.exports = { processGmailOrders, reprocessTodayOrders, backfillPositions };
