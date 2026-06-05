const { google } = require('googleapis');
const fs = require('fs');

let _cachedClient = null;

/**
 * Создаёт и кэширует Google Auth-клиент.
 * Повторные вызовы возвращают тот же объект — без повторной аутентификации.
 */
async function getAuthClient() {
  if (_cachedClient) return _cachedClient;

  // --- Режим 1: OAuth2 refresh_token ---
  if (process.env.GMAIL_REFRESH_TOKEN) {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GMAIL_CLIENT_ID,
      process.env.GMAIL_CLIENT_SECRET,
      'https://developers.google.com/oauthplayground'
    );
    oauth2Client.setCredentials({ refresh_token: process.env.GMAIL_REFRESH_TOKEN });
    _cachedClient = oauth2Client;
    return _cachedClient;
  }

  // --- Режим 2: Service Account ---
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    let credentials;
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
      clientOptions: process.env.GMAIL_IMPERSONATE_EMAIL
        ? { subject: process.env.GMAIL_IMPERSONATE_EMAIL }
        : undefined,
    });
    _cachedClient = await auth.getClient();
    return _cachedClient;
  }

  throw new Error('Не найдены credentials. Укажите GMAIL_REFRESH_TOKEN или GOOGLE_SERVICE_ACCOUNT_JSON в .env');
}

function getGmailClient(auth) {
  return google.gmail({ version: 'v1', auth });
}

function getSheetsClient(auth) {
  return google.sheets({ version: 'v4', auth });
}

module.exports = { getAuthClient, getGmailClient, getSheetsClient };
