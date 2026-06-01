/**
 * Shared date parsing utilities.
 * parseDateStr avoids new Date() to prevent UTC timezone shift.
 */

/**
 * Parses date string without new Date() to avoid UTC shift.
 * Supported formats:
 *   "2026-05-31"          ISO prefix (from Date.toISOString)
 *   "2026-05-31 11:50:00" ISO-like with time
 *   "31.05.2026"          Russian
 *   "31.05.2026, 11:50"   Russian with time (from toLocaleString ru-RU)
 *   "31.05.26"            Russian short year
 * Returns { dd, mm, yyyy } or null.
 */
function parseDateStr(s) {
  if (!s) return null;
  s = s.toString().trim();

  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return { yyyy: iso[1], mm: iso[2], dd: iso[3] };

  const ru = s.match(/^(\d{2})\.(\d{2})\.(\d{2,4})/);
  if (ru) {
    const yyyy = ru[3].length === 2 ? `20${ru[3]}` : ru[3];
    return { yyyy, mm: ru[2], dd: ru[1] };
  }

  return null;
}

/** Returns today's date as YYYYMMDD string in local (server) time */
function todayDateKey() {
  const t = new Date();
  const dd   = String(t.getDate()).padStart(2, '0');
  const mm   = String(t.getMonth() + 1).padStart(2, '0');
  const yyyy = String(t.getFullYear());
  return `${yyyy}${mm}${dd}`;
}

/** Returns today as { yyyy, mm, dd } in local time */
function todayParts() {
  const t = new Date();
  return {
    yyyy: String(t.getFullYear()),
    mm:   String(t.getMonth() + 1).padStart(2, '0'),
    dd:   String(t.getDate()).padStart(2, '0'),
  };
}

module.exports = { parseDateStr, todayDateKey, todayParts };
