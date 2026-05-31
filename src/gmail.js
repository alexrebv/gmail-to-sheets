const { getGmailClient, getAuthClient } = require('./auth');
const { appendRowsToSheet, ensureSheetExists } = require('./sheets');

const LABEL_NAME = process.env.GMAIL_LABEL || 'Transfer';
const SEARCH_QUERY_BASE = process.env.GMAIL_QUERY || 'subject:"отправлен сотрудником"';
const AFTER_DATE = process.env.GMAIL_AFTER || '2026/05/30';
const SHEET_NAME = process.env.SHEET_NAME || 'Отправлен';

/**
 * Парсит тему письма, возвращает { object, orderNumber, orderDate }
 * Формат темы: "Заказ для ресторана <Объект> #123-456-789 создан DD/MM/YY HH:MM"
 */
function parseSubject(subject) {
  const regex = /Заказ для ресторана (.+?) (#[\d-]+) создан (\d{2}\/\d{2}\/\d{2} \d{2}:\d{2})/;
  const match = subject.match(regex);
  if (match) {
    return {
      object: match[1].trim(),
      orderNumber: match[2].trim(),
      orderDate: match[3].trim(),
    };
  }
  return { object: '', orderNumber: '', orderDate: '' };
}

/**
 * Декодирует base64url в строку
 */
function decodeBody(data) {
  if (!data) return '';
  const base64 = data.replace(/-/g, '+').replace(/_/g, '/');
  return Buffer.from(base64, 'base64').toString('utf-8');
}

/**
 * Извлекает plain-text тело письма из структуры payload
 */
function extractPlainBody(payload) {
  if (!payload) return '';

  // Прямое тело без вложений
  if (payload.mimeType === 'text/plain' && payload.body?.data) {
    return decodeBody(payload.body.data);
  }

  // Многочастное письмо — рекурсивно ищем text/plain
  if (payload.parts) {
    for (const part of payload.parts) {
      const text = extractPlainBody(part);
      if (text) return text;
    }
  }

  return '';
}

/**
 * Получает или создаёт Gmail-лейбл по имени, возвращает его id
 */
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

/**
 * Основная функция обработки писем
 */
async function processGmailOrders() {
  try {
    const auth = await getAuthClient();
    const gmail = getGmailClient(auth);

    const labelId = await getOrCreateLabel(gmail, LABEL_NAME);
    const query = `${SEARCH_QUERY_BASE} after:${AFTER_DATE} -label:${LABEL_NAME}`;

    console.log(`Поиск писем: ${query}`);

    // Получаем все сообщения (с пагинацией)
    const messageIds = [];
    let pageToken;
    do {
      const res = await gmail.users.messages.list({
        userId: 'me',
        q: query,
        maxResults: 100,
        pageToken,
      });
      const msgs = res.data.messages || [];
      messageIds.push(...msgs.map(m => m.id));
      pageToken = res.data.nextPageToken;
    } while (pageToken);

    console.log(`Найдено писем: ${messageIds.length}`);

    if (messageIds.length === 0) {
      console.log('Нет новых писем.');
      return;
    }

    // Читаем каждое письмо
    const newRows = [];
    const processedIds = [];

    for (const id of messageIds) {
      const res = await gmail.users.messages.get({
        userId: 'me',
        id,
        format: 'full',
      });

      const msg = res.data;
      const headers = msg.payload?.headers || [];

      const dateHeader = headers.find(h => h.name === 'Date')?.value || '';
      const subject = headers.find(h => h.name === 'Subject')?.value || '';

      const emailDate = dateHeader ? new Date(dateHeader) : new Date(parseInt(msg.internalDate));
      const body = extractPlainBody(msg.payload);
      const { object, orderNumber, orderDate } = parseSubject(subject);

      // Добавляем строку даже если regex не совпал — всё равно фиксируем письмо
      newRows.push([
        emailDate,
        object || subject,  // если не распарсилось — пишем тему целиком
        orderNumber,
        orderDate,
        body.substring(0, 500), // первые 500 символов тела (можно убрать или расширить)
      ]);

      processedIds.push(id);
    }

    // Записываем в таблицу
    const HEADERS = ['Дата письма', 'Объект', 'Номер заказа', 'Дата заказа', 'Тело письма'];
    await ensureSheetExists(SHEET_NAME, HEADERS);
    await appendRowsToSheet(SHEET_NAME, newRows);

    console.log(`Записано строк: ${newRows.length}`);

    // Помечаем письма лейблом Transfer
    for (const id of processedIds) {
      await gmail.users.messages.modify({
        userId: 'me',
        id,
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
