/**
 * index.js — точка входа Railway-сервиса
 *
 * Запускает:
 *  1. channelBot  — webhook-сервер для @AcceptODChannel (мгновенно пишет в «Принят»)
 *  2. Gmail reader      — cron, читает почту → «Отправлен»
 *  3. Send orders       — cron, шлёт заказы в Telegram
 *  4. Check status      — cron, сверяет «Принят»/«Вычерк» с «Отправлен»
 */

require('dotenv').config();
const cron = require('node-cron');
const { processGmailOrders }         = require('./gmail');
const { sendOrdersToTelegram }       = require('./sendOrders');
const { updateOrderStatusAndNotify } = require('./checkStatus');
const { startChannelBot, sendEndDaySummary } = require('./channelBot');
const { getConfig }                          = require('./config');

const DEFAULT_CRON_GMAIL       = '*/15 * * * *';
const DEFAULT_CRON_SEND_ORDERS = '0 8 * * *';
const DEFAULT_CRON_STATUS      = '0 6,10,14,18,22 * * *';
const DEFAULT_CRON_END_DAY     = '30 22 * * *';

async function start() {
  console.log(`[${ts()}] ═══ Gmail → Sheets worker запущен ═══`);

  // 1. Запускаем webhook-бот немедленно
  await startChannelBot();

  // 2. Загружаем расписания из таблицы
  let cfg = {};
  try {
    cfg = await getConfig();
  } catch (e) {
    console.warn(`[${ts()}] Настройки недоступны, использую дефолты: ${e.message}`);
  }

  const cronGmail      = cfg.CRON_GMAIL       || DEFAULT_CRON_GMAIL;
  const cronSendOrders = cfg.CRON_SEND_ORDERS || DEFAULT_CRON_SEND_ORDERS;
  const cronStatus     = cfg.CRON_STATUS      || DEFAULT_CRON_STATUS;
  const cronEndDay     = cfg.CRON_END_DAY     || DEFAULT_CRON_END_DAY;

  const TZ = cfg.TIMEZONE || 'Europe/Moscow';
  const cronOpts = { timezone: TZ };

  console.log(`  Gmail reader   : ${cronGmail} (${TZ})`);
  console.log(`  Send orders TG : ${cronSendOrders}`);
  console.log(`  Check status   : ${cronStatus}`);
  console.log(`  End day summary: ${cronEndDay}`);

  // 3. Первый запуск Gmail reader сразу
  run('Gmail reader', processGmailOrders);

  // 4. Cron-задачи
  cron.schedule(cronGmail,      () => run('Gmail reader',           processGmailOrders),  cronOpts);
  cron.schedule(cronSendOrders, () => run('Send orders → Telegram', sendOrdersToTelegram), cronOpts);
  cron.schedule(cronStatus,     () => run('Check status + notify',  updateOrderStatusAndNotify), cronOpts);
  cron.schedule(cronEndDay,     () => run('End day summary',        runEndDaySummary), cronOpts);
}

async function runEndDaySummary() {
  const cfg = await getConfig();
  if (!cfg.TELEGRAM_TOKEN || !cfg.TELEGRAM_CHAT_ID) {
    console.warn('[end_day] TELEGRAM_TOKEN или TELEGRAM_CHAT_ID не заданы');
    return;
  }
  await sendEndDaySummary(cfg.TELEGRAM_CHAT_ID, cfg.TELEGRAM_THREAD_ID || null, cfg);
}

async function run(label, fn) {
  console.log(`\n[${ts()}] ▶ ${label}`);
  try {
    await fn();
    console.log(`[${ts()}] ✓ ${label} завершён`);
  } catch (err) {
    console.error(`[${ts()}] ✗ ${label} — ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  }
}

function ts() {
  return new Date().toISOString().replace('T', ' ').slice(0, 19);
}

start();
