const { google } = require('googleapis');
const fs = require('fs');

/**
 * Создаёт OAuth2-клиент для Gmail (от имени пользователя через refresh_token)
 * или Service Account с domain-wide delegation.
 *
 * Поддерживаемые режимы (задаются через env):
 *  1. GMAIL_REFRESH_TOKEN + GMAIL_CLIENT_ID + GMAIL_CLIENT_SECRET  → OAuth2 (рекомендуется для личного Gmail)
 *  2. GOOGLE_SERVICE_ACCOUNT_JSON + GMAIL_IMPERSONATE_EMAIL         → Service Account + delegation (для Workspace)
 */
async function getAuthClient() {
  // --- Режим 1: OAuth2 refresh_token ---
  if (process.env.GMAIL_REFRESH_TOKEN) {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GMAIL_CLIENT_ID,
      process.env.GMAIL_CLIENT_SECRET,
      'https://developers.google.com/oauthplayground'
    );

    oauth2Client.setCredentials({
      refresh_token: process.env.GMAIL_REFRESH_TOKEN,
    });

    return oauth2Client;
  }

  // --- Режим 2: Service Account с domain-wide delegation ---
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    let credentials;

    // Поддержка: путь к файлу ИЛИ JSON-строка прямо в env
    if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON.trim().startsWith('{')) {
      credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    } else {
      credentials = JSON.parse(fs.readFileSync(process.env.GOOGLE_SERVICE_ACCOUNT_JSON, 'utf8'));
    }

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        'https://www.googleapis.com/auth/gmail.modify',
        'https://www.googleapis.com/auth/gmail.labels',
        'https://www.googleapis.com/auth/spreadsheets',
      ],
      // Обязательно для доступа к Gmail через SA
      clientOptions: process.env.GMAIL_IMPERSONATE_EMAIL
        ? { subject: process.env.GMAIL_IMPERSONATE_EMAIL }
        : undefined,
    });

    return auth.getClient();
  }

  throw new Error(
    'Не найдены credentials. Укажите GMAIL_REFRESH_TOKEN или GOOGLE_SERVICE_ACCOUNT_JSON в .env'
  );
}

function getGmailClient(auth) {
  return google.gmail({ version: 'v1', auth });
}

function getSheetsClient(auth) {
  return google.sheets({ version: 'v4', auth });
}

module.exports = { getAuthClient, getGmailClient, getSheetsClient };
