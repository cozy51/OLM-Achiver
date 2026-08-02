// Injected into the Outlook tab with chrome.scripting.executeScript, so this
// function is serialised with toString(): it must stay self-contained and must
// not reference imports or closure variables. It only collects raw candidates —
// choosing between them is sender.js's job, which keeps the untestable part
// small.
export function collectSenderCandidates() {
  const MAX_CANDIDATES = 400;
  const textOf = element => element?.textContent?.trim() || "";
  const firstOf = (selectors, root = document) => {
    for (const selector of selectors) {
      const found = root.querySelector(selector);
      if (found) return found;
    }
    return null;
  };

  const readingPane = firstOf([
    '[data-app-section="MailReadCompose"]',
    '[aria-label*="閲覧ウィンドウ"]',
    '[aria-label*="Reading pane"]'
  ]);
  const root = readingPane || document;
  const body = firstOf([
    '[data-testid="message-body"]',
    '[aria-label="メッセージ本文"]',
    '[aria-label="Message body"]',
    '[role="document"]'
  ], root);
  if (!body) {
    return { ok: false, error: "開いているメール本文を見つけられませんでした。メールを1通開いてから再実行してください。" };
  }

  const bodyTop = body.getBoundingClientRect().top;
  const candidates = [];
  const push = (source, raw, rect, order, persona) => {
    if (candidates.length >= MAX_CANDIDATES) return;
    const value = String(raw || "").trim();
    if (!value || value.length > 200) return;
    candidates.push({
      source,
      raw: value,
      order,
      persona: Boolean(persona),
      top: rect ? rect.top : null,
      bottom: rect ? rect.bottom : null,
      left: rect ? rect.left : null
    });
  };

  // 1. Attributes Outlook has used to mark the sender across its builds.
  for (const selector of [
    '[data-testid="message-sender"]',
    '[data-testid="sender-name"]',
    '[data-testid="SenderPersona"]',
    '[automation-id="SenderPersona"]',
    '[aria-label^="差出人"]',
    '[aria-label^="送信者"]',
    '[aria-label^="From"]',
    '[aria-label^="Sender"]'
  ]) {
    const element = root.querySelector(selector);
    if (!element) continue;
    const rect = element.getBoundingClientRect();
    push("explicit", element.getAttribute("aria-label"), rect, 0);
    push("explicit", element.getAttribute("title"), rect, 0);
    push("explicit", textOf(element), rect, 0);
  }

  // 2. The selected mail list row. Its accessible name begins with the sender,
  // and it is the same row the message URL is taken from, so it cannot drift to
  // a different message than the one being saved.
  const selectedRow = document.querySelector(
    '[role="option"][aria-selected="true"], [aria-selected="true"], [data-is-selected="true"]'
  );
  if (selectedRow) {
    const rect = selectedRow.getBoundingClientRect();
    const label = selectedRow.getAttribute("aria-label") || "";
    label.split(/[、,]\s*/).slice(0, 4).forEach((segment, index) => push("selected-row", segment, rect, index));
    for (const element of selectedRow.querySelectorAll("[title]")) {
      push("selected-row", element.getAttribute("title"), element.getBoundingClientRect(), 0);
    }
  }

  // 3. The message header, bounded below by the "宛先:" row so recipients and
  // the body cannot be mistaken for the sender.
  const recipientRows = [];
  for (const element of root.querySelectorAll("span, div, [aria-label]")) {
    const rect = element.getBoundingClientRect();
    if (!rect.width || !rect.height || rect.top >= bodyTop || rect.height > 60) continue;
    const label = element.getAttribute("aria-label") || textOf(element);
    if (/^\s*(?:宛先|To|Cc|CC|BCC)\s*[:：]/i.test(label)) {
      // Prefer a small, leaf-like label: a large wrapper can start above the
      // sender and would then cut it out of the search area entirely.
      recipientRows.push({ top: rect.top, length: label.length, area: rect.width * rect.height });
    }
  }
  recipientRows.sort((a, b) => a.length - b.length || a.area - b.area || b.top - a.top);
  const recipientsTop = recipientRows.length ? recipientRows[0].top : bodyTop;

  // Application chrome sits above the message too. Excluding it by landmark is
  // reliable, unlike matching the wording of every command Outlook can show.
  const CHROME = '[role="banner"], [role="toolbar"], [role="menubar"], [role="navigation"], [role="tablist"], [role="search"], header, nav';
  const isAppChrome = element => Boolean(element.closest(CHROME));
  // Outlook opens a profile card from the sender, so a popup owner near the
  // header is a far better cue than the element's wording.
  const isPersona = element => Boolean(
    element.closest('[data-testid*="persona" i], [automation-id*="persona" i], [class*="persona" i]') ||
    element.getAttribute("aria-haspopup") ||
    element.closest('a[href^="mailto:"]')
  );

  for (const element of root.querySelectorAll('[title], [aria-label], a[href^="mailto:"], button, [role="button"], span')) {
    const rect = element.getBoundingClientRect();
    if (!rect.width || !rect.height || rect.bottom > recipientsTop + 1) continue;
    if (isAppChrome(element)) continue;
    const persona = isPersona(element);
    push("header", element.getAttribute("title"), rect, 0, persona);
    push("header", element.getAttribute("aria-label"), rect, 0, persona);
    push("header", textOf(element), rect, 0, persona);
  }

  return {
    ok: true,
    candidates,
    recipientsTop,
    bodyTop,
    // Reported so a failure can be told apart from an empty result: without a
    // reading pane or a "宛先:" row the header search area is only a guess.
    readingPane: Boolean(readingPane),
    recipientRowFound: recipientRows.length > 0
  };
}
