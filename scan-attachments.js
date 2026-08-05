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
  // Outlook ids the attachment well as "MSG_<item>_ATTACHMENTS" and hangs its
  // test hooks off the same word, so the area is found by what Outlook calls it
  // rather than by where it happens to be drawn.
  const NAMED_WELL = [
    '[id$="_ATTACHMENTS"]',
    '[data-testid*="attachment" i]',
    '[automation-id*="attachment" i]'
  ];
  // Outlook also labels the list inside the well "添付ファイル"/"画像添付ファイル",
  // which is how an image attachment is found at all. But a row in the message
  // list says "添付ファイルあり" in its own label, so this wording only means the
  // well while it is read inside the reading pane.
  const LABELLED_WELL = [
    '[aria-label*="添付ファイル"]',
    '[aria-label*="Attachment" i]'
  ];
  const findWells = (scope, selectors) => {
    const matched = [...scope.querySelectorAll(selectors.join(", "))];
    // The well and the list inside it both match, and each is scanned in full,
    // so the inner one is dropped rather than walked a second time.
    return matched.filter(well => !matched.some(other => other !== well && other.contains(well)));
  };
  const inPane = readingPane ? findWells(readingPane, [...NAMED_WELL, ...LABELLED_WELL]) : [];
  // A build that draws the well outside the pane would otherwise hide every
  // attachment, so the document is still searched — but only for a well that
  // names itself, and never inside a row of the message list. Six such rows,
  // each labelled "添付ファイルあり" for a mail of its own, were once listed in a
  // PDF as the attachments of the open mail, which had none: out here a
  // [role="option"] is another message, not a file.
  const outside = findWells(document, NAMED_WELL).filter(well => !well.closest('[role="option"]'));
  const wells = inPane.length ? inPane : outside;

  // An image attachment is drawn as a thumbnail with no visible name, so the
  // name can be anywhere the accessibility tree can reach: the tile's label,
  // the alt text of the preview, the name a download anchor would save under,
  // or a hidden element the tile points at.
  const NAMED = ["aria-label", "title", "alt", "download", "name"];
  const labelledBy = element => (element.getAttribute("aria-labelledby") || "")
    .split(/\s+/)
    .filter(Boolean)
    .map(id => document.getElementById(id)?.textContent?.trim() || "");

  // Outlook draws each attachment as one [role="option"], with the file name on
  // the option and the preview on a [draggable="true"] child. Resolving to the
  // option first — rather than to whichever ancestor happens to be nearest —
  // keeps a tile and its own contents from being counted as two files.
  const tiles = new Map();
  const tileOf = element => {
    const found = element.closest('[role="option"], [role="listitem"]') ||
      element.closest('a[href], [draggable="true"]');
    // Whether a tile was found at all is reported: wording loose in the well —
    // the summary line, the buttons under it — is not an attachment, and only a
    // tile can vouch for a name that has no extension to prove it is one.
    const tile = found || element;
    if (!tiles.has(tile)) tiles.set(tile, tiles.size);
    return { tile, index: tiles.get(tile), inTile: Boolean(found) };
  };

  const entries = [];
  for (const well of wells) {
    for (const element of [well, ...well.querySelectorAll("*")]) {
      if (entries.length >= MAX_ENTRIES) break;
      const labels = [
        ...NAMED.map(name => element.getAttribute(name)),
        ...labelledBy(element),
        // Only leaves: a wrapper's text runs the tile's fields together.
        element.children.length ? "" : (element.textContent || "").trim(),
        // Builds that keep the name in a data attribute are covered too: a
        // value is only used when it reads as a file name, so collecting them
        // costs nothing.
        ...[...element.attributes].filter(item => item.name.startsWith("data-")).map(item => item.value)
      ].filter(label => label && label.length <= MAX_LABEL);
      if (!labels.length) continue;
      const { tile, index, inTile } = tileOf(element);
      // A cloud attachment is a real link; a file stored in the message is a
      // tile that downloads through script, so it has no URL of its own.
      const anchor = tile.matches("a[href]") ? tile : tile.querySelector("a[href]");
      // Whether the label belongs to the tile or to something drawn inside it.
      // A tile's own label is the attachment; its parts include buttons.
      entries.push({ tile: index, own: element === tile, inTile, labels, href: anchor?.href || "" });
    }
  }
  // Reported so "no attachments" can be told apart from "the well was found but
  // nothing in it reads as a file name", which is a result the user must see.
  return { entries, pageUrl: location.href, wells: wells.length };
}
