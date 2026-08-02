// Turns the raw labels collected from the Outlook attachment well into the file
// list written into the PDF. Kept apart from the injected scan so the fragile
// part stays a dumb collector and every judgement can be unit tested without a
// browser.

// Outlook joins the fields of an attachment tile into one label, as in
// "画像添付ファイル、写真.jpg、6 MB". Only one of those fields is the file name,
// and an extension is what tells it apart from the size or the tile's category.
// Requiring one also keeps the well's own summary line ("添付ファイル 2 個 (6 MB)")
// and its buttons out of the list.
const FILE_NAME = /^.{1,120}\.[A-Za-z0-9]{1,8}$/;
// A size written without a space ("6.5MB") has the shape of a file name, so it
// is the one form that has to be named to be excluded.
const FILE_SIZE = /^[\d.,]+\s*(?:B|KB|MB|GB|TB|バイト)$/i;
// A message with more tiles than this is not what the list is for; the cap only
// stops a mislabelled page from filling the PDF.
const MAX_ATTACHMENTS = 40;

export function attachmentNameFrom(value) {
  const text = String(value || "")
    // Outlook draws the file-type icon with a private-use glyph in the label.
    .replace(/[\uE000-\uF8FF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "";
  // The whole label is tried first: a file name may itself contain a comma.
  for (const segment of [text, ...text.split(/[、,]\s*/)]) {
    const name = segment
      .replace(/^\s*(?:添付ファイル|Attachment)\s*[:：]?\s*/i, "")
      // A trailing "(6 MB)" is the size Outlook appends after the name.
      .replace(/\s*[(（][^()（）]*[)）]\s*$/, "")
      .trim();
    if (FILE_NAME.test(name) && !FILE_SIZE.test(name)) return name;
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
    const existing = found.get(name);
    if (existing) {
      // The name usually sits on an outer element and the link on an inner one.
      if (!existing.url && url) existing.url = url;
      continue;
    }
    if (found.size >= MAX_ATTACHMENTS) break;
    found.set(name, { name, url });
  }
  return [...found.values()];
}
