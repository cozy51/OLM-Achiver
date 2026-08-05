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
// An Outlook item — a mail attached to a mail — carries a subject where a file
// carries a name, so there is no extension to cut at. Outlook writes the kind on
// the tile instead, and that is what makes the subject safe to read as a name:
// without it, any wording in the well would qualify.
const ITEM_KIND = /Outlook\s*(?:アイテム|item)|メール\s*メッセージ|mail\s*message/i;
// The commands Outlook draws inside a tile. They are never the attachment. Only
// a label that does not name its kind is judged this way — a subject may well
// contain "保存", and the kind already proves such a label is the attachment.
const TILE_COMMAND = /開く|ダウンロード|保存|プレビュー|削除|選択|その他|オプション|メニュー|^(?:Open|Download|Save|Preview|Delete|Select|More|Options|Menu)\b/i;
// A message with more tiles than this is not what the list is for; the cap only
// stops a mislabelled page from filling the PDF.
const MAX_ATTACHMENTS = 40;
// An address has the shape of a file name: the ".nishi" of
// "yasuhiro.nishi@akl.muratec.co.jp" reads as an extension, and the name would
// be the person. Outlook writes an address wherever the well carries a person —
// the sender of an attached mail, the alt text of a contact photo — so the span
// an address covers is worked out first and every extension inside it is passed
// over. The domain deliberately stops short of requiring a dot: in
// "meeting@2026.pdf" only "meeting@2026" is read as an address, which leaves the
// file its name.
const ADDRESS = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+/g;
// A label that is an address and nothing else names nobody's file, whichever
// route reads it.
const ADDRESS_ONLY = /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+$/;

const clean = value => String(value || "")
  // Outlook draws the file-type icon with a private-use glyph in the label.
  .replace(/[\uE000-\uF8FF]/g, " ")
  .replace(/\s+/g, " ")
  .trim();

const usableName = name =>
  Boolean(name) && name.length <= 120 && !FILE_SIZE.test(name) && !/[\\/]/.test(name) &&
  !ADDRESS_ONLY.test(name);


export function attachmentNameFrom(value) {
  const text = clean(value);
  for (const field of text.split(FIELDS)) {
    const addresses = [...field.matchAll(ADDRESS)].map(match => [match.index, match.index + match[0].length]);
    for (const match of field.matchAll(EXTENSION)) {
      if (!HAS_LETTER.test(match[0])) continue;
      const end = match.index + match[0].length;
      if (addresses.some(([from, to]) => match.index < to && end > from)) continue;
      // Everything before the extension belongs to the name: it can hold spaces
      // and brackets, as "Image (1).jpg" does, so it cannot be tokenised. An
      // address in front of it is not part of it, though — the name starts
      // after the last one, so a person cc'd beside the file stays out of it.
      const start = addresses.filter(([, to]) => to <= match.index).map(([, to]) => to).pop() || 0;
      const name = field.slice(start, end).replace(PREFIX, "").trim();
      if (!usableName(name)) continue;
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

// The subject of a mail attached to a mail. It is only ever read from a label
// that names the kind, so the boundary is the kind itself: what precedes it is
// the subject, what follows is the size and the commands.
export function attachmentSubjectFrom(value) {
  const text = clean(value);
  const kind = text.match(ITEM_KIND);
  if (!kind) return "";
  const subject = text.slice(0, kind.index).replace(/[、,]\s*$/, "").trim();
  return usableName(subject) ? subject : "";
}

// Taken from a real read view: an Outlook item's tile is labelled
// "【79584】FST PBN CST LP & CST見学報告 開く 250 KB" — the kind is drawn beside
// the tile but kept out of its label, so the command is the only boundary left.
// The same shape carries the file name of an ordinary attachment.
const COMMAND_TOKEN = /\s(?:開く|プレビュー|ダウンロード|保存|削除|選択|その他|オプション|メニュー|Open|Preview|Download|Save|Delete|Select|More|Options|Menu)(?=\s|$)/;

export function attachmentNameBeforeCommand(value) {
  const text = clean(value);
  const command = text.match(COMMAND_TOKEN);
  if (!command) return "";
  const name = text.slice(0, command.index).trim();
  return usableName(name) && /[\p{L}\p{N}]/u.test(name) ? name : "";
}

// A last resort for a tile whose label holds the name alone, with nothing in it
// to mark where the name ends because there is nothing after it.
function looseName(value) {
  const text = clean(value);
  if (!usableName(text) || ITEM_KIND.test(text) || TILE_COMMAND.test(text)) return "";
  return /[\p{L}\p{N}]/u.test(text) ? text : "";
}

function nameOfTile(entries) {
  const labelsOf = only => entries
    .filter(entry => (only ? entry.own : true))
    .flatMap(entry => entry.labels || []);
  // A file name is the strongest evidence and needs no help from the tile.
  const named = labelsOf(false).map(attachmentNameFrom).find(Boolean);
  if (named) return named;
  const subject = labelsOf(false).map(attachmentSubjectFrom).find(Boolean);
  if (subject) return subject;
  // Past this point the name has no mark of its own, so the tile has to vouch
  // for it: a [role="option"] in the attachment well is an attachment, and the
  // only question left is where its label stops naming it. Wording loose in the
  // well — the summary line, the buttons under it — vouches for nothing.
  if (!entries.some(entry => entry.inTile)) return "";
  // The tile's own label names the attachment; its parts include buttons.
  return labelsOf(true).map(attachmentNameBeforeCommand).find(Boolean) ||
    labelsOf(true).map(looseName).find(Boolean) ||
    labelsOf(false).map(looseName).find(Boolean) || "";
}

export function listAttachments(scan) {
  // One tile is one attachment. Grouping by tile rather than by name keeps the
  // label and the image inside a tile from counting twice, and still lists two
  // tiles that happen to carry the same file name as the two files they are.
  const tiles = new Map();
  for (const entry of scan?.entries || []) {
    const key = Number.isInteger(entry?.tile) ? `tile:${entry.tile}` : `entry:${tiles.size}`;
    if (!tiles.has(key)) tiles.set(key, []);
    tiles.get(key).push(entry);
  }

  const found = [];
  for (const entries of tiles.values()) {
    if (found.length >= MAX_ATTACHMENTS) break;
    const name = nameOfTile(entries);
    if (!name) continue;
    // The name usually sits on an outer element and the link on an inner one.
    const url = entries.map(entry => attachmentLink(entry.href, scan?.pageUrl)).find(Boolean) || "";
    found.push({ name, url });
  }
  return found;
}
