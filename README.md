# Gmail → Google Sheets

Сервис читает входящие письма Gmail по критериям, парсит данные заказов и записывает их в Google Таблицу. Запускается на Railway.app по расписанию (cron).

---

## Структура проекта

```
src/
  index.js   — точка входа, cron-расписание
  gmail.js   — чтение и парсинг писем
  sheets.js  — запись в Google Sheets
  auth.js    — авторизация (OAuth2 / Service Account)
```

---

## Настройка

### 1. Google Cloud Console

1. Создайте проект на [console.cloud.google.com](https://console.cloud.google.com)
2. Включите **Gmail API** и **Google Sheets API**
3. Выберите способ авторизации:

---

#### Вариант A — OAuth2 (личный Gmail, рекомендуется)

1. Создайте `OAuth 2.0 Client ID` → тип **Web application**
2. В `Authorized redirect URIs` добавьте: `https://developers.google.com/oauthplayground`
3. Получите `refresh_token` через [OAuth Playground](https://developers.google.com/oauthplayground):
   - Шестерёнка → поставьте галку "Use your own OAuth credentials", вставьте Client ID и Secret
   - Выберите скоупы: `https://www.googleapis.com/auth/gmail.modify` и `https://www.googleapis.com/auth/spreadsheets`
   - Нажмите "Authorize APIs" → "Exchange authorization code for tokens"
   - Скопируйте `refresh_token`

Переменные окружения:
```
GMAIL_CLIENT_ID=...
GMAIL_CLIENT_SECRET=...
GMAIL_REFRESH_TOKEN=...
```

---

#### Вариант B — Service Account (Google Workspace)

1. Создайте Service Account → скачайте JSON-ключ
2. В Admin Console настройте **Domain-wide delegation** для этого SA
3. Скоупы для delegation:
   - `https://www.googleapis.com/auth/gmail.modify`
   - `https://www.googleapis.com/auth/spreadsheets`

Переменные окружения:
```
GOOGLE_SERVICE_ACCOUNT_JSON={"type":"service_account",...}  # вставьте весь JSON как строку
GMAIL_IMPERSONATE_EMAIL=user@yourdomain.com
```

> **Важно**: SA не имеет доступа к Gmail напрямую без domain-wide delegation. Для обычного личного Gmail используйте Вариант A.

---

### 2. Google Sheets

1. Создайте таблицу
2. Дайте доступ:
   - **Вариант A**: никаких дополнительных действий (токен уже от вашего аккаунта)
   - **Вариант B**: добавьте email сервисного аккаунта как редактора таблицы
3. Скопируйте ID таблицы из URL

```
SPREADSHEET_ID=1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms
```

---

### 3. Переменные окружения

Скопируйте `.env.example` в `.env` и заполните:

```bash
cp .env.example .env
```

| Переменная | Описание | По умолчанию |
|---|---|---|
| `SPREADSHEET_ID` | ID Google Таблицы | **обязательно** |
| `GMAIL_LABEL` | Лейбл для помечания обработанных писем | `Transfer` |
| `GMAIL_QUERY` | Поисковый запрос Gmail | `subject:"отправлен сотрудником"` |
| `GMAIL_AFTER` | Обрабатывать письма после этой даты | `2026/05/30` |
| `SHEET_NAME` | Название листа в таблице | `Отправлен` |
| `CRON_SCHEDULE` | Расписание запуска (cron-формат) | `*/15 * * * *` |
| `TIMEZONE` | Таймзона для отображения дат | `Europe/Moscow` |

---

### 4. Деплой на Railway

```bash
# Инициализируйте git
git init && git add . && git commit -m "init"

# Залейте на GitHub и подключите к Railway
# Или используйте Railway CLI:
railway login
railway init
railway up
```

В Railway добавьте все переменные из `.env` через раздел **Variables**.

---

## Формат данных в таблице

| Дата письма | Объект | Номер заказа | Дата заказа | Тело письма |
|---|---|---|---|---|
| 31.05.2026 10:23 | Ресторан Центр | #123-456-789 | 30/05/26 09:00 | ... |

- Парсинг темы: `Заказ для ресторана <Объект> #NNN-NNN-NNN создан DD/MM/YY HH:MM`
- Если тема не совпала с шаблоном — в колонку "Объект" записывается тема целиком
- Тело письма: первые 500 символов (можно изменить в `gmail.js`)
- Повторная обработка исключается через лейбл `Transfer`

---

## Локальный запуск

```bash
npm install
node src/index.js
```
