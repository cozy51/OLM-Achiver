// Turns the raw labels collected from the Outlook attachment well into the file
// list written into the PDF. Kept apart from the injected scan so the fragile
// part stays a dumb collector and every judgement can be unit tested without a
// browser.

// A tile's accessible name is the file name followed by what can be done with
// it and how big it is — Outlook writes it as "Image (1).jpg 開く 6 MB", with no
// separator to key on. The extension is the fixed point: the name ends there,
// and what follows describes the tile rather than the file. Requiring one also
// keeps the well's summary line ("添付ファイル 2 個 (6 MB)") and its buttons out.
const EXTENSION = /\.[A-Za-z0-9]{1,8}(?![\p{L}\p{N}])/gu;
// ".5" in "1.5 MB" has the shape of an extension; a real one carries a letter.
const HAS_LETTER = /[A-Za-z]/;
// A size written without a space ("6.5MB") passes that test, so it is the one
// form that has to be named to be excluded.
const FILE_SIZE = /^[\d.,]+\s*(?:B|KB|MB|GB|TB|バイト)$/i;
// Outlook does join fields with a Japanese comma elsewhere. An ASCII comma is
// left alone: it appears inside file names far more often than between fields.
const FIELDS = /[、]\s*/;
// What Outlook can put in front of the name. English words are not stripped —
// "Image" and "Open" begin real file names as readily as they label a tile.
const PREFIX = /^(?:(?:画像)?添付ファイル|Attachment)\s*[:：]?\s*|^(?:開く|プレビュー|ダウンロード|選択)\s+/i;
// A message with more tiles than this is not what the list is for; the cap only
// stops a mislabelled page from filling the PDF.
const MAX_ATTACHMENTS = 40;

export function attachmentNameFrom(value) {
  const text = String(value || "")
    // Outlook draws the file-type icon with a private-use glyph in the label.
    .replace(/[\uE000-\uF8FF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  for (const field of text.split(FIELDS)) {
    for (const match of field.matchAll(EXTENSION)) {
      if (!HAS_LETTER.test(match[0])) continue;
      // Everything before the extension belongs to the name: it can hold spaces
      // and brackets, as "Image (1).jpg" does, so it cannot be tokenised.
      const name = field.slice(0, match.index + match[0].length).replace(PREFIX, "").trim();
      if (!name || name.length > 120 || FILE_SIZE.test(name)) continue;
      // A file name carries no path separator; a URL in a data attribute does.
      if (/[\\/]/.test(name)) continue;
      // An extension with nothing in front of it is not a file name.
      if (name.length <= match[0].length) continue;
      return name;
    }
  }
  return "";
}

export function attachmentLink(href, pageUrl) {
  const value = String(href || "");
  if (!/^https?:/i.test(value)) return "";
  // A tile whose anchor only re-points at the open message is not a file link.
  return value.split("#")[0] === String(pageUrl || "").split("#")[0] ? "" : value;
}

export function listAttachments(scan) {
  const found = new Map();
  for (const entry of scan?.entries || []) {
    const name = (entry?.labels || []).map(attachmentNameFrom).find(Boolean);
    if (!name) continue;
    const url = attachmentLink(entry.href, scan?.pageUrl);
    // One tile is one file. Grouping by tile rather than by name keeps the
    // label and the image inside a tile from counting twice, and still lists
    // two tiles that happen to carry the same file name as the two files they
    // are. An entry outside any tile falls back to its name.
    const key = Number.isInteger(entry.tile) ? `tile:${entry.tile}` : `name:${name}`;
    const existing = found.get(key);
    if (existing) {
      // The name usually sits on an outer element and the link on an inner one.
      if (!existing.url && url) existing.url = url;
      continue;
    }
    if (found.size >= MAX_ATTACHMENTS) break;
    found.set(key, { name, url });
  }
  return [...found.values()];
}
