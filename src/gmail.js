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
 * Парсит тему письма — поддерживает форматы:
 *   #20260-749-0006   (цифры и дефисы)
 *   #2026-ОД15-3944   (с кириллицей)
 * Время может быть 1:23 или 01:23
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
 * Извлекает поставщика из HTML-тела письма iiko.
 * Поставщик находится в ячейках class="column5..." таблицы заказа.
 * Пример: "ИП Григорян - 1уп 8шт, заказ в уп" → "ИП Григорян"
 */
function extractSupplierFromHtml(html) {
  if (!html) return '';
  const tdRegex = /class="column5[^"]*"[^>]*>([\s\S]*?)<\/td>/gi;
  let match;
  while ((match = tdRegex.exec(html)) !== null) {
    const text = match[1]
      .replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
      .trim();
    if (text && text.length > 2) {
      const dashIdx = text.indexOf(' - ');
      return dashIdx > 0 ? text.substring(0, dashIdx).trim() : text.trim();
    }
  }
  return '';
}

/** Fallback: поиск поставщика в plain-text теле */
function extractSupplierFromPlain(text) {
  if (!text) return '';
  const lines = text.split('\n');
  for (const line of lines) {
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
        emailDate,                   // A — Дата письма
        object || subject,           // B — Объект (без ФГ)
        orderNumber,                 // C — Номер заказа
        orderDate,                   // D — Дата заказа
        supplier,                    // E — Поставщик
        emailDate,                   // F — Дата отправки
        '',                          // G — Юр.лицо
        '',                          // H — Направлено
        plainBody.substring(0, 300) || htmlBody.substring(0, 300), // I — Тело
      ]);

      processedIds.push(id);
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
