// Scoring and clean-up for the sender candidates collected from the Outlook
// DOM. Kept apart from the injected scan so the fragile part stays a dumb
// collector and every judgement can be unit tested without a browser.

const EMAIL = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/;
const DATE_LIKE = /\d{4}\s*(?:年|[-/.])\s*\d{1,2}\s*(?:月|[-/.])\s*\d{1,2}|\d{1,2}\s*[:：]\s*\d{2}/;
// Outlook renders commands and column labels next to the sender. None of them
// ever name a person, so they must never win the ranking.
const NOT_A_PERSON = new RegExp([
  "^(?:宛先|To|Cc|CC|BCC|差出人|From|件名|Subject|返信|全員に返信|転送|Reply|Reply all|Forward",
  "|その他|More|削除|Delete|アーカイブ|Archive|フラグ|Flag|移動|Move|印刷|Print|閉じる|Close",
  "|開封済み|未読|Unread|Read|Copilot|下書き|Draft|添付ファイル|Attachment|重要|Important",
  // A reply prefix means the value is a subject line, not a person.
  "|RE|FW|FWD)(?:\\s*[:：]|\\s*$)",
  "|(?:Ctrl|Alt|Shift)\\s*\\+|^\\d+$|件のメッセージ|^\\s*$"
].join(""), "i");
// Outlook writes colleagues as "山田 太郎 (Yamada,T.)"; the Japanese UI uses
// full-width brackets and commas for the same form.
const NAMED_PERSON = /[(（][^()（）]*[,，][^()（）]*[)）]/;
// Below this a candidate is a stray label rather than a plausible person, so it
// is not even offered as a choice.
const MIN_SENDER_SCORE = 3;
// A candidate needs at least one strong person-like signal before the popup
// fills the field on its own; anything weaker is offered as a choice instead.
export const CONFIDENT_SENDER_SCORE = 9;

export function extractEmailAddress(value) {
  return String(value || "").match(EMAIL)?.[0] || "";
}

export function cleanSenderName(value) {
  return String(value || "")
    // Outlook draws its icons with private-use glyphs inside the same label.
    .replace(/[\uE000-\uF8FF]/g, " ")
    .replace(/^\s*(?:差出人|送信者|送信元|発信者|発信元|From|Sender)\s*[:：]?\s*/i, "")
    .replace(/[<（(]\s*[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\s*[>）)]?/g, " ")
    .replace(/[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/g, " ")
    .replace(/\s*(?:プロフィール\s?カードを開く|さんのプロフィール|Open profile card|profile card)\S*/gi, " ")
    // A persona label appends further fields after the display name, as in
    // "原田 清枝（Harada,K.）, 送信済み". The name ends at its own bracket.
    .replace(/^([^()（）]*[(（][^()（）]*[)）])\s*[,、].*$/, "$1")
    .replace(/[;；,、]+\s*$/, "")
    .replace(/\s+/g, " ")
    .trim();
}

export function isSenderLike(name) {
  const value = String(name || "").trim();
  if (!value || value.length > 80) return false;
  if (NOT_A_PERSON.test(value)) return false;
  if (DATE_LIKE.test(value)) return false;
  return /[\p{L}\p{N}]/u.test(value);
}

export function senderKey(value) {
  return String(value || "").toLowerCase().replace(/[\s　.,、。()（）<>]/g, "");
}

export function scoreSenderCandidate(candidate, context = {}) {
  const raw = String(candidate?.raw || "");
  const name = cleanSenderName(raw);
  const address = extractEmailAddress(raw);
  if (!name && !address) return null;
  if (name && !isSenderLike(name)) return null;

  // Everything in the header that is not a person still reads like plain text
  // ("More items", "アプリ"), so wording alone cannot separate them. Require
  // positive evidence of a person instead of listing what to reject: an
  // address, Outlook's "氏名 (Surname,I.)" form, or a profile-card element.
  const personEvidence = Boolean(address) ||
    NAMED_PERSON.test(name) ||
    candidate.persona ||
    /プロフィール|profile card/i.test(raw);
  if (candidate.source === "header" && !personEvidence) return null;

  let score = 0;
  if (candidate.source === "explicit") score += 8;
  // The mail list row starts with the sender, so only its first field is a
  // strong signal; later fields are the subject and the preview text.
  if (candidate.source === "selected-row") score += candidate.order ? 1 : 4;
  if (candidate.source === "header") score += 2;
  if (address) score += 5;
  // Outlook shows colleagues as "山田 太郎 (Yamada,T.)".
  if (NAMED_PERSON.test(name)) score += 3;
  if (candidate.persona || /プロフィール|profile card/i.test(raw)) score += 3;
  if (candidate.source === "header" &&
      Number.isFinite(context.recipientsTop) && Number.isFinite(candidate.bottom)) {
    // The sender sits on the row directly above "宛先:".
    const gap = context.recipientsTop - candidate.bottom;
    if (gap >= -1 && gap < 40) score += 3;
    else if (gap >= -1 && gap < 90) score += 1;
  }
  return { name, address, score, source: candidate.source };
}

export function rankSenders(scan) {
  const merged = new Map();
  for (const candidate of scan?.candidates || []) {
    const entry = scoreSenderCandidate(candidate, scan);
    if (!entry) continue;
    const key = senderKey(entry.name || entry.address);
    const existing = merged.get(key);
    if (!existing) {
      merged.set(key, { ...entry, sources: new Set([entry.source]) });
      continue;
    }
    existing.score = Math.max(existing.score, entry.score);
    existing.name = existing.name || entry.name;
    existing.address = existing.address || entry.address;
    existing.sources.add(entry.source);
  }
  return [...merged.values()]
    .map(entry => ({
      name: entry.name || entry.address,
      address: entry.address,
      // Independent parts of the page naming the same person is the strongest
      // evidence available, and it is what single-signal scraping lacked.
      score: entry.score + (entry.sources.size - 1) * 6,
      sources: [...entry.sources].sort()
    }))
    .filter(entry => entry.score >= MIN_SENDER_SCORE)
    .sort((a, b) => b.score - a.score || a.name.localeCompare(b.name));
}
