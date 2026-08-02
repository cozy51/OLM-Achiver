export function stripReplyPrefix(subject) {
  let value = (subject || "").trim();
  const prefix = /^(?:(?:re|fw|fwd|返信|転送)\s*(?:\[\d+\])?\s*[:：]\s*)/i;
  while (prefix.test(value)) value = value.replace(prefix, "").trim();
  return value || "件名なし";
}

export function sanitizeFilenamePart(value, fallback) {
  const cleaned = (value || "")
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
    .replace(/[. ]+$/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return (cleaned || fallback).slice(0, 80);
}

export function buildFilename({ sender, date, subject }) {
  const instant = parseMailDate(date);
  const parts = instant ? new Intl.DateTimeFormat("ja-JP", {
      year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", hourCycle: "h23"
    }).formatToParts(instant).reduce((all, item) => ({ ...all, [item.type]: item.value }), {}) : null;
  const day = parts ? `${parts.year}${parts.month}${parts.day}` : "受信日不明";
  const time = parts ? `${parts.hour}${parts.minute}` : "時刻不明";
  return `📩_${sanitizeFilenamePart(sender, "差出人不明")}_${day}_${time}_${sanitizeFilenamePart(stripReplyPrefix(subject), "件名なし")}.pdf`;
}

export function parseMailDate(value) {
  if (value instanceof Date) return Number.isNaN(value.valueOf()) ? null : value;
  const source = String(value || "").trim();
  if (!source) return null;
  // Outlook's Japanese UI uses forms such as
  // "2026-07-13 (月) 20:41" and "2026年7月13日 20:41". Constructing
  // the Date from numeric parts avoids browser-dependent parsing.
  const match = source.match(/(\d{4})\s*(?:年|[-/.])\s*(\d{1,2})\s*(?:月|[-/.])\s*(\d{1,2})\s*日?(?:\s*\([^)]*\))?\s*T?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
  if (match) {
    const [, year, month, day, hour, minute] = match.map(Number);
    const parsed = new Date(year, month - 1, day, hour, minute);
    if (parsed.getFullYear() === year && parsed.getMonth() === month - 1 && parsed.getDate() === day) return parsed;
    return null;
  }
  const parsed = new Date(source);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

export function isSupportedOutlookUrl(value) {
  try {
    const url = new URL(value);
    return url.protocol === "https:" && [
      "outlook.office.com",
      "outlook.office365.com",
      "outlook.cloud.microsoft"
    ].includes(url.hostname);
  } catch {
    return false;
  }
}

export function selectOutlookTab(tabs) {
  const candidates = tabs.filter(tab => tab.id && isSupportedOutlookUrl(tab.url));
  return candidates.find(tab => tab.active) ||
    candidates.sort((a, b) => (b.lastAccessed || 0) - (a.lastAccessed || 0))[0];
}
