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
process.env.TZ = process.env.TIMEZONE || 'Europe/Moscow';
const cron = require('node-cron');
const db = require('./db');
const { processGmailOrders, reprocessTodayOrders, backfillPositions } = require('./gmail');
const { sendOrdersToTelegram }                        = require('./sendOrders');
const { updateOrderStatus, updateOrderStatusAndNotify } = require('./checkStatus');
const { startChannelBot, sendTodayOrders }            = require('./channelBot');
const { getConfig }                                   = require('./config');

const DEFAULT_CRON_GMAIL       = '* * * * *';
const DEFAULT_CRON_SEND_ORDERS = '0 8 * * *';
const DEFAULT_CRON_STATUS      = '0 6,10,14,18,22 * * *';
const DEFAULT_CRON_END_DAY     = '30 22 * * *';
const DEFAULT_CRON_BUY         = '0 12,13,14,15,16 * * *';
const DEFAULT_CRON_REPROCESS   = '*/5 * * * *';

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
  const cronBuy        = cfg.CRON_BUY         || DEFAULT_CRON_BUY;
  const cronReprocess  = cfg.CRON_REPROCESS   || DEFAULT_CRON_REPROCESS;

  console.log(`  Gmail reader   : ${cronGmail}`);
  console.log(`  Send orders TG : ${cronSendOrders}`);
  console.log(`  Check status   : ${cronStatus}`);
  console.log(`  End of day     : ${cronEndDay}`);
  console.log(`  Today orders   : ${cronBuy}`);
  console.log(`  Reprocess GO   : ${cronReprocess}`);

  // 3. Разовый бэкфилл листа «Позиции» — включается вручную переменной
  // окружения BACKFILL_POSITIONS=true, после того как отработает — можно убрать.
  if (process.env.BACKFILL_POSITIONS === 'true') {
    await run('Backfill Позиции', backfillPositions);
  }

  // 4. Первый запуск Gmail reader сразу
  run('Gmail reader', processGmailOrders);

  // 5. Cron-задачи
  cron.schedule('* * * * *',    () => run('Check status (update)', updateOrderStatus));
  cron.schedule(cronGmail,      () => run('Gmail reader',           processGmailOrders));
  cron.schedule(cronStatus,     () => run('Check status + notify',  updateOrderStatusAndNotify));
  cron.schedule(cronEndDay,     () => run('End of day report',       runEndDay));
  cron.schedule(cronBuy,        () => run('Today orders → Telegram', runTodayOrders));
  cron.schedule(cronReprocess,  () => run('Reprocess today GO',      reprocessTodayOrders));

  // Схема Postgres — ПОСЛЕ регистрации расписания и без await: если база
  // недоступна и подключение зависнет, это не должно помешать cron-задачам
  // (иначе молча пропадают отчёты по накладным). Схема всё равно готовится
  // лениво при первой записи в базу.
  db.init().catch(e => console.error('[db] init:', e.message));
}

async function runEndDay() {
  const cfg      = await getConfig();
  const chatId   = cfg.TELEGRAM_CHAT_ID   || process.env.TELEGRAM_CHAT_ID;
  const threadId = cfg.TELEGRAM_THREAD_ID || process.env.TELEGRAM_THREAD_ID || null;
  if (!chatId) { console.warn('[runEndDay] TELEGRAM_CHAT_ID не задан'); return; }
  const { sendEndOfDayReport } = require('./channelBot');
  await sendEndOfDayReport(chatId, threadId, cfg);
}

async function runTodayOrders() {
  const cfg      = await getConfig();
  const chatId   = cfg.TELEGRAM_CHAT_ID   || process.env.TELEGRAM_CHAT_ID;
  const threadId = cfg.TELEGRAM_THREAD_ID || process.env.TELEGRAM_THREAD_ID || null;
  if (!chatId) { console.warn('[runTodayOrders] TELEGRAM_CHAT_ID не задан'); return; }
  await sendTodayOrders(chatId, threadId, cfg);
}

// Задачи, которые сейчас выполняются. При минутном расписании очередной запуск
// может прийти раньше, чем закончился предыдущий (Gmail + Sheets + Telegram
// вполне укладываются в минуту не всегда) — накладывающиеся запуски приводили бы
// к повторной обработке одних и тех же писем. Поэтому пропускаем тик, если
// предыдущий запуск этой же задачи ещё идёт.
// label -> { token, startedAt }. Токен нужен, чтобы «зависший» прогон,
// завершившись позже, не снял отметку у актуального.
const _running = new Map();
// Если задача идёт дольше этого времени, считаем её застрявшей и запускаем
// заново: иначе один подвисший прогон навсегда глушил бы расписание.
const STALE_MS = 30 * 60 * 1000;
let _runToken = 0;

async function run(label, fn) {
  const cur = _running.get(label);
  if (cur) {
    const age = Date.now() - cur.startedAt;
    if (age < STALE_MS) {
      console.log(`[${ts()}] ⏭ ${label} — предыдущий запуск ещё идёт, пропускаем тик`);
      return;
    }
    console.warn(`[${ts()}] ⚠ ${label} — предыдущий запуск идёт ${Math.round(age / 60000)} мин, запускаю заново`);
  }
  const token = ++_runToken;
  _running.set(label, { token, startedAt: Date.now() });
  console.log(`\n[${ts()}] ▶ ${label}`);
  try {
    await fn();
    console.log(`[${ts()}] ✓ ${label} завершён`);
  } catch (err) {
    console.error(`[${ts()}] ✗ ${label} — ошибка: ${err.message}`);
    if (err.stack) console.error(err.stack);
  } finally {
    // Снимаем отметку только если она наша (иначе затрём актуальный прогон).
    if (_running.get(label)?.token === token) _running.delete(label);
  }
}

// Время в логах — в рабочем часовом поясе, том же, по которому срабатывает
// cron. toISOString() всегда отдаёт UTC, из-за чего строка вида
// «[2026-08-14 03:00:00] Check status» соответствовала запуску в 06:00 МСК и
// выглядела как сбитое время.
function ts() {
  return new Date().toLocaleString('sv-SE', {
    timeZone: process.env.TIMEZONE || 'Europe/Moscow',
  });
}

start();
