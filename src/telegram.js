/**
 * telegram.js
 * Отправка сообщений в Telegram с поддержкой топиков (message_thread_id).
 */

const https = require('https');

async function sendMessage(token, chatId, text, threadId = null, waitMs = 5000) {
  const payload = {
    chat_id: chatId,
    text,
    parse_mode: 'Markdown',
  };

  // threadId передаём если это число > 0
  const tid = parseInt(threadId);
  if (!isNaN(tid) && tid > 0) payload.message_thread_id = tid;

  await new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const req = https.request(
      {
        hostname: 'api.telegram.org',
        path: `/bot${token}/sendMessage`,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (chunk) => (data += chunk));
        res.on('end', () => {
          const parsed = JSON.parse(data);
          if (!parsed.ok) console.error('[Telegram] Ошибка:', parsed.description, '| thread_id:', tid);
          resolve(parsed);
        });
      }
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });

  if (waitMs > 0) await sleep(waitMs);
}

async function sendLongMessage(token, chatId, text, threadId = null, maxLen = 4000, waitMs = 5000) {
  const parts = splitMessage(text, maxLen);
  for (const part of parts) {
    await sendMessage(token, chatId, part, threadId, waitMs);
  }
}

function splitMessage(text, maxLen) {
  const parts = [];
  while (text.length > 0) {
    if (text.length <= maxLen) { parts.push(text); break; }
    let slice = text.slice(0, maxLen);
    const lastNl = slice.lastIndexOf('\n');
    if (lastNl > 0) slice = slice.slice(0, lastNl + 1);
    parts.push(slice);
    text = text.slice(slice.length).trim();
  }
  return parts;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

module.exports = { sendMessage, sendLongMessage, sleep };
