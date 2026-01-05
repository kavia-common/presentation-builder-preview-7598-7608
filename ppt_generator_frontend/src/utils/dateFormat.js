/**
 * Date formatting helpers shared by preview and PPT export.
 */

/**
 * PUBLIC_INTERFACE
 * Format an ISO date string (YYYY-MM-DD from <input type="date">) into "DD MMM YYYY".
 *
 * Example: "2026-01-05" -> "05 Jan 2026"
 *
 * Returns empty string when input is empty/invalid.
 */
export function formatDdMmmYyyy(dateStr) {
  /** This is a public function. */
  if (dateStr == null) return '';
  const s = String(dateStr).trim();
  if (!s) return '';

  // Expecting YYYY-MM-DD (native date input output). Parse as UTC to avoid timezone shifts.
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) return '';

  const year = Number(m[1]);
  const month = Number(m[2]); // 1..12
  const day = Number(m[3]);

  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return '';
  if (month < 1 || month > 12) return '';
  if (day < 1 || day > 31) return '';

  const dt = new Date(Date.UTC(year, month - 1, day));
  if (Number.isNaN(dt.getTime())) return '';

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const dd = String(day).padStart(2, '0');
  const mmm = months[month - 1];

  return `${dd} ${mmm} ${year}`;
}
