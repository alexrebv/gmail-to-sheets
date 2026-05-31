/**
 * index.js — точка входа Railway-сервиса
 *
 * Три задачи по cron-расписанию (читаются из листа «Настройки»):
 *   CRON_GMAIL        — читает Gmail и пишет в «Отправлен»        (по умолчанию каждые 15 мин)
 *   CRON_SEND_ORDERS  — отправляет новые заказы в Telegram        (по умолчанию в 08:00)
 *   CRON_STATUS       — проверяет статусы Принят/Вычерк + уведом. (по умолчанию в 09:00)
 */

require('dotenv').config();
const cron = require('node-cron');
const { processGmailOrders }           = require('./gmail');
const { sendOrdersToTelegram }         = require('./sendOrders');
const { updateOrderStatusAndNotify }   = require('./checkStatus');
const { getConfig }                    = require('./config');

const DEFAULT_CRON_GMAIL       = '*/15 * * * *';
const DEFAULT_CRON_SEND_ORDERS = '0 8 * * *';
const DEFAULT_CRON_STATUS      = '0 9 * * *';

async function start() {
  console.log(`[${ts()}] ═══ Gmail → Sheets worker запущен ═══`);

  let cfg = {};
  try {
    cfg = await getConfig();
  } catch (e) {
    console.warn(`[${ts()}] Не удалось загрузить Настройки, использую дефолты: ${e.message}`);
  }

  const cronGmail      = cfg.CRON_GMAIL       || DEFAULT_CRON_GMAIL;
  const cronSendOrders = cfg.CRON_SEND_ORDERS || DEFAULT_CRON_SEND_ORDERS;
  const cronStatus     = cfg.CRON_STATUS      || DEFAULT_CRON_STATUS;

  console.log(`  Gmail reader   : ${cronGmail}`);
  console.log(`  Send orders TG : ${cronSendOrders}`);
  console.log(`  Check status   : ${cronStatus}`);

  // ── Запуск при старте ────────────────────────────────────────────────────
  run('Gmail reader',     processGmailOrders);
  // sendOrders и checkStatus запускаем только по расписанию, не сразу

  // ── Cron ─────────────────────────────────────────────────────────────────
  cron.schedule(cronGmail, () =>
    run('Gmail reader', processGmailOrders));

  cron.schedule(cronSendOrders, () =>
    run('Send orders → Telegram', sendOrdersToTelegram));

  cron.schedule(cronStatus, () =>
    run('Check status + notify', updateOrderStatusAndNotify));
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
