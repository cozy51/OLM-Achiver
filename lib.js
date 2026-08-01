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
  const instant = date instanceof Date && !Number.isNaN(date.valueOf()) ? date : new Date();
  const parts = new Intl.DateTimeFormat("ja-JP", {
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", hourCycle: "h23"
  }).formatToParts(instant).reduce((all, item) => ({ ...all, [item.type]: item.value }), {});
  const day = `${parts.year}${parts.month}${parts.day}`;
  const time = `${parts.hour}${parts.minute}`;
  return `📩_${sanitizeFilenamePart(sender, "差出人不明")}_${day}_${time}_${sanitizeFilenamePart(stripReplyPrefix(subject), "件名なし")}.pdf`;
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
