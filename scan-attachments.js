// Injected into the Outlook tab with chrome.scripting.executeScript, so this
// function is serialised with toString(): it must stay self-contained and must
// not reference imports or closure variables. It only collects raw labels and
// links — deciding which label is a file name is attachments.js's job, which
// keeps the untestable part small.
export function collectAttachmentTiles() {
  const MAX_ENTRIES = 600;
  const MAX_LABEL = 200;
  const firstOf = selectors => {
    for (const selector of selectors) {
      const found = document.querySelector(selector);
      if (found) return found;
    }
    return null;
  };

  const readingPane = firstOf([
    '[data-app-section="ItemContainer"]',
    '[data-app-section="MailReadCompose"]',
    '[aria-label*="閲覧ウィンドウ"]',
    '[aria-label*="Reading pane"]'
  ]);
  const root = readingPane || document;

  // Outlook ids the attachment well as "MSG_<item>_ATTACHMENTS" and labels the
  // list inside it "添付ファイル"/"画像添付ファイル", so the area is found by what
  // Outlook calls it rather than by where it happens to be drawn.
  const matched = [...root.querySelectorAll([
    '[id$="_ATTACHMENTS"]',
    '[data-testid*="attachment" i]',
    '[automation-id*="attachment" i]',
    '[aria-label*="添付ファイル"]',
    '[aria-label*="Attachment" i]'
  ].join(", "))];
  // The well and the list inside it both match, and each is scanned in full, so
  // the inner one is dropped rather than walked a second time.
  const wells = matched.filter(well => !matched.some(other => other !== well && other.contains(well)));

  const entries = [];
  for (const well of wells) {
    for (const element of [well, ...well.querySelectorAll("*")]) {
      if (entries.length >= MAX_ENTRIES) break;
      const labels = [
        element.getAttribute("aria-label"),
        element.getAttribute("title"),
        // Only leaves: a wrapper's text runs the tile's fields together.
        element.children.length ? "" : (element.textContent || "").trim()
      ].filter(label => label && label.length <= MAX_LABEL);
      if (!labels.length) continue;
      // A cloud attachment is a real link; a file stored in the message is a
      // tile that downloads through script, so it has no URL of its own.
      const tile = element.closest('a[href], [role="option"], [role="listitem"], [draggable="true"]') || element;
      const anchor = tile.matches("a[href]") ? tile : tile.querySelector("a[href]");
      entries.push({ labels, href: anchor?.href || "" });
    }
  }
  return { entries, pageUrl: location.href };
}
