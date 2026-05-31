const { getGmailClient, getAuthClient } = require('./auth');
const { appendRowsToSheet, ensureSheetExists } = require('./sheets');
const { getConfig } = require('./config');

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

function extractByMime(payload, mimeType) {
  if (!payload) return '';
  if (payload.mimeType === mimeType && payload.body?.data) {
    return decodeBody(payload.body.data);
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
      .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')
      .trim();
    if (text && text.length > 2) return text;
  }

  // Fallback: ищем любую ячейку column0 с ИП/ООО/АО
  const cellRegex = /class="column0[^"]*"[^>]*>([\s\S]*?)<\/td>/gi;
  let m;
  while ((m = cellRegex.exec(html)) !== null) {
    const text = m[1].replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
    if (/^(ИП|ООО|АО|ЗАО|ПАО)\s+/i.test(text)) return text;
  }

  return '';
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

    const labelId = await getOrCreateLabel(gmail, LABEL_NAME);
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

    for (const id of messageIds) {
      const res = await gmail.users.messages.get({ userId: 'me', id, format: 'full' });
      const msg     = res.data;
      const headers = msg.payload?.headers || [];

      const dateHeader = headers.find(h => h.name === 'Date')?.value || '';
      const subject    = headers.find(h => h.name === 'Subject')?.value || '';
      const emailDate  = dateHeader ? new Date(dateHeader) : new Date(parseInt(msg.internalDate));

      const htmlBody  = extractByMime(msg.payload, 'text/html');
      const plainBody = extractByMime(msg.payload, 'text/plain');

      const { object, orderNumber, orderDate } = parseSubject(subject);
      const supplier = extractSupplierFromHtml(htmlBody) || extractSupplierFromPlain(plainBody);

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

      processedIds.push(id);
      console.log(`  ✓ ${object} | ${orderNumber} | ${supplier}`);
    }

    const HEADERS = [
      'Дата письма', 'Объект', 'Номер заказа', 'Дата заказа',
      'Поставщик', 'Дата отправки', 'Юр.лицо', 'Направлено', 'Тело письма',
    ];
    await ensureSheetExists(SHEET_NAME, HEADERS);
    await appendRowsToSheet(SHEET_NAME, newRows);
    console.log(`Записано строк: ${newRows.length}`);

    for (const id of processedIds) {
      await gmail.users.messages.modify({
        userId: 'me', id,
        requestBody: { addLabelIds: [labelId] },
      });
    }
    console.log(`Помечено писем лейблом "${LABEL_NAME}": ${processedIds.length}`);

  } catch (err) {
    console.error(`[processGmailOrders] Ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
}

module.exports = { processGmailOrders };
