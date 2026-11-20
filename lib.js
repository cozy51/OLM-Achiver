export function sanitizeFilenamePart(value, fallback) {
  const cleaned = (value || "")
    // "_" joins the fields of the file name, so a sender written as
    // "Ito Takeshi / 伊藤剛" must not read as another field boundary.
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "-")
    .replace(/[. ]+$/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return (cleaned || fallback).slice(0, 80);
}

export function buildFilename({ sender, date, attachments, comment }) {
  const instant = parseMailDate(date);
  const parts = instant ? new Intl.DateTimeFormat("ja-JP", {
      year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", hourCycle: "h23"
    }).formatToParts(instant).reduce((all, item) => ({ ...all, [item.type]: item.value }), {}) : null;
  const day = parts ? `${parts.year}${parts.month}${parts.day}` : "受信日不明";
  const time = parts ? `${parts.hour}${parts.minute}` : "時刻不明";
  // The clip marks a mail that carried files, so a folder listing answers "was
  // anything attached to this one" without opening the PDF. It reflects what
  // the list in the PDF holds: a mail whose attachment area could not be read
  // is not marked, and the popup says so at save time.
  const mark = (attachments || []).length ? "📩📎" : "📩";
  // An optional note from the popup sits right after the mark, in braces, so
  // it sorts and reads apart from the sender. Without a note the name is
  // exactly what it was before the field existed.
  const note = sanitizeFilenamePart(comment, "");
  // The name ends with the received time. A subject made long names that
  // crowded the folder listing; the subject is still the heading of the PDF
  // and the Markdown, so it is not lost.
  return `${mark}${note ? `{${note}}` : ""}_${sanitizeFilenamePart(sender, "差出人不明")}_${day}_${time}.pdf`;
}

// The PDF carries the full attachment list, so the popup only has to make the
// result checkable at a glance: how many files were found, and enough names to
// tell that the right message was saved.
export function describeAttachments(attachments, diagnostics = {}) {
  const names = (attachments || []).map(item => String(item?.name || "").trim()).filter(Boolean);
  // Names the file the list went into: the PDF, or the Markdown when that is
  // the only thing saved.
  const doc = diagnostics.document || "PDF";
  if (!names.length) {
    // Saying nothing is right for a message with nothing attached. An
    // attachment area that was found but could not be read is a different
    // result, and reporting it as silence would leave the PDF quietly short of
    // what the user asked it to record.
    return diagnostics.wells
      ? `添付欄は見つかりましたが、ファイル名を読み取れませんでした。${doc}に一覧は入っていません。`
      : "";
  }
  const rest = names.length > 3 ? ` ほか${names.length - 3}件` : "";
  return `添付ファイル ${names.length}件の名前とリンクを${doc}に記載しました: ${names.slice(0, 3).join(" / ")}${rest}`;
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

// What the popup saves: the PDF, the Markdown, or both. Anything else falls
// back to the PDF alone, which is what the extension did before the choice.
export const SAVE_FORMATS = ["pdf", "markdown", "both"];

export function normalizeSaveFormat(value) {
  return SAVE_FORMATS.includes(value) ? value : "pdf";
}

// The choice used to be a single "also save Markdown" checkbox. A user who had
// ticked it keeps getting both files instead of silently losing the Markdown.
export function storedSaveFormat(stored = {}) {
  if (SAVE_FORMATS.includes(stored.saveFormat)) return stored.saveFormat;
  return stored.saveMarkdown === true ? "both" : "pdf";
}
