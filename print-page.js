export function outlookPrintMode(action, sourceUrl) {
  const WRAPPER_ID = "olm-pdf-document";
  const STYLE_ID = "olm-pdf-style";

  if (action === "print") {
    const suggestedName = sourceUrl || document.title;
    document.title = suggestedName.replace(/\.pdf$/i, "");
    const finish = () => {
      document.getElementById(WRAPPER_ID)?.remove();
      document.getElementById(STYLE_ID)?.remove();
      if (globalThis.__olmPreviousTitle !== undefined) {
        document.title = globalThis.__olmPreviousTitle;
        delete globalThis.__olmPreviousTitle;
      }
    };
    addEventListener("afterprint", finish, { once: true });
    // Let executeScript return before the modal print preview opens.
    setTimeout(() => window.print(), 0);
    return { ok: true };
  }

  if (action === "end") {
    document.getElementById(WRAPPER_ID)?.remove();
    document.getElementById(STYLE_ID)?.remove();
    if (globalThis.__olmPreviousTitle !== undefined) {
      document.title = globalThis.__olmPreviousTitle;
      delete globalThis.__olmPreviousTitle;
    }
    return { ok: true };
  }

  const elementText = element => element?.textContent?.trim() || "";
  const first = (selectors, root = document) => {
    for (const selector of selectors) {
      const found = root.querySelector(selector);
      if (found && (elementText(found) || found.getAttribute("aria-label") || found.getAttribute("title"))) return found;
    }
    return null;
  };

  const isMessageUrl = value => {
    try {
      const url = new URL(value, location.origin);
      return url.origin === location.origin && (
        /\/mail\/deeplink\/read\//i.test(url.pathname) ||
        /\/mail\/(?:[^/]+\/)?id\/[^/]+/i.test(url.pathname)
      );
    } catch {
      return false;
    }
  };

  const findMessageUrl = currentUrl => {
    if (isMessageUrl(currentUrl)) return new URL(currentUrl, location.origin).href;

    // Outlook may keep the address bar at /mail/ while a message is open in
    // the reading pane. Prefer links inside the selected/current list row.
    const selectors = [
      '[aria-selected="true"] a[href]',
      '[data-is-selected="true"] a[href]',
      '[aria-current="true"][href]',
      'a[aria-current="true"][href]',
      'a[href*="/deeplink/read/"]',
      'a[href*="/mail/"][href*="/id/"]'
    ];
    for (const selector of selectors) {
      for (const link of document.querySelectorAll(selector)) {
        const href = link.href || link.getAttribute("href");
        if (isMessageUrl(href)) return new URL(href, location.origin).href;
      }
    }

    // Some Outlook builds store the selected item id as data rather than an
    // anchor. Convert that id to Outlook's stable /mail/inbox/id/... route.
    const selected = document.querySelector('[aria-selected="true"], [data-is-selected="true"]');
    const itemHolder = selected?.closest('[data-item-id], [data-itemid]') ||
      selected?.querySelector('[data-item-id], [data-itemid]');
    const itemId = itemHolder?.getAttribute("data-item-id") || itemHolder?.getAttribute("data-itemid");
    if (itemId) {
      let decodedId = itemId;
      try { decodedId = decodeURIComponent(itemId); } catch { /* already raw */ }
      return `${location.origin}/mail/inbox/id/${encodeURIComponent(decodedId)}`;
    }
    return "";
  };

  const findReceivedDate = (searchRoot, body) => {
    const bodyTop = body.getBoundingClientRect().top;
    const datePattern = /\d{4}\s*(?:年|[-/.])\s*\d{1,2}\s*(?:月|[-/.])\s*\d{1,2}\s*日?(?:\s*\([^)]*\))?\s*T?\s*\d{1,2}\s*[:：]\s*\d{2}/;
    const candidates = [];
    for (const element of searchRoot.querySelectorAll('time, [data-testid*="date" i], [data-testid*="time" i], [title], [aria-label], span')) {
      const rect = element.getBoundingClientRect();
      if (!rect.width || !rect.height || rect.top >= bodyTop) continue;
      const values = [
        element.getAttribute("datetime"),
        element.getAttribute("title"),
        element.getAttribute("aria-label"),
        elementText(element)
      ];
      const value = values.find(candidate => datePattern.test(candidate || ""));
      if (!value) continue;
      candidates.push({ value: value.match(datePattern)?.[0] || value, top: rect.top, left: rect.left });
    }
    // The received timestamp is normally at the right edge of the top header.
    candidates.sort((a, b) => a.top - b.top || b.left - a.left);
    return candidates[0]?.value || "";
  };

  const readingPane = first([
    '[data-app-section="MailReadCompose"]',
    '[aria-label*="閲覧ウィンドウ"]',
    '[aria-label*="Reading pane"]'
  ]);
  const root = readingPane || document;
  const body = first([
    '[data-testid="message-body"]',
    '[aria-label="メッセージ本文"]',
    '[aria-label="Message body"]',
    '[role="document"]'
  ], root);
  if (!body) return { ok: false, error: "開いているメール本文を見つけられませんでした。メールを1通開いてから再実行してください。" };

  const container = body.closest('[data-testid="message-item"], [role="listitem"], article') || body.parentElement || body;
  const subject = elementText(first([
    '[data-testid="message-subject"]', 'h1[title]',
    '[role="heading"][aria-level="2"]', '[role="heading"]'
  ], root)) || document.title.replace(/^メール\s*-\s*/i, "");
  const receivedDate = findReceivedDate(root, body);
  const messageUrl = findMessageUrl(sourceUrl);
  if (!messageUrl) {
    return { ok: false, error: "開いているメールの個別URLを取得できませんでした。メール一覧で対象メールを選択してから再実行してください。" };
  }

  document.getElementById(WRAPPER_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
  globalThis.__olmPreviousTitle = document.title;
  const wrapper = document.createElement("main");
  wrapper.id = WRAPPER_ID;
  const link = document.createElement("a");
  link.className = "olm-source-link";
  link.href = messageUrl;
  link.textContent = `元のメールを Outlook で開く: ${messageUrl}`;
  const heading = document.createElement("h1");
  heading.textContent = subject;
  wrapper.append(link, heading, container.cloneNode(true));
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = `
    @page { size: A4; margin: 14mm; }
    body > * { display: none !important; }
    body > #${WRAPPER_ID} { display: block !important; color: #1b1b1b; font: 12px/1.55 Arial, "Yu Gothic", sans-serif; }
    #${WRAPPER_ID} * { max-width: 100% !important; }
    #${WRAPPER_ID} h1 { font-size: 18px; border-bottom: 1px solid #bbb; padding-bottom: 8px; }
    #${WRAPPER_ID} .olm-source-link { display: block; padding: 10px; margin-bottom: 14px; border: 1px solid #3574c5; background: #eef6ff; color: #0645ad; overflow-wrap: anywhere; }
    #${WRAPPER_ID} button, #${WRAPPER_ID} [role="button"] { display: none !important; }
  `;
  document.head.appendChild(style);
  document.body.appendChild(wrapper);
  document.title = subject;
  return {
    ok: true,
    metadata: {
      subject,
      date: receivedDate
    }
  };
}

export function buildPrintPageExpression(action, sourceUrl = "") {
  return `(${outlookPrintMode.toString()})(${JSON.stringify(action)}, ${JSON.stringify(sourceUrl)})`;
}
