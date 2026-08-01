(function () {
  // executeScript() may be used as a fallback for a tab that predates the
  // extension installation. Do not register a second message listener when
  // the manifest content script is already active.
  if (globalThis.__olmPdfArchiverLoaded) return;
  globalThis.__olmPdfArchiverLoaded = true;
  const MESSAGE = "olm-pdf-print-mode";
  let cleanup = null;

  const text = (element) => element?.textContent?.trim() || "";
  const first = (selectors, root = document) => {
    for (const selector of selectors) {
      const found = root.querySelector(selector);
      if (found && text(found)) return found;
    }
    return null;
  };

  function findMessage() {
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
    if (!body) throw new Error("開いているメール本文を見つけられませんでした。メールを1通開いてから再実行してください。");

    const container = body.closest('[data-testid="message-item"], [role="listitem"], article') || body.parentElement;
    const subject = text(first([
      '[data-testid="message-subject"]',
      'h1[title]',
      '[role="heading"][aria-level="2"]',
      '[role="heading"]'
    ], root)) || document.title.replace(/^メール\s*-\s*/i, "");
    const senderElement = first([
      '[data-testid="message-sender"]',
      '[data-testid="sender-name"]',
      '[aria-label^="差出人"]',
      '[aria-label^="From"]'
    ], container || root);
    const timeElement = first([
      'time[datetime]',
      '[data-testid="message-date"]',
      '[title*="202"]'
    ], container || root);
    return {
      body,
      container: container || body,
      subject,
      sender: text(senderElement).replace(/^(差出人|From)\s*[:：]?\s*/i, "") || "差出人不明",
      date: timeElement?.getAttribute("datetime") || timeElement?.getAttribute("title") || text(timeElement) || new Date().toISOString()
    };
  }

  function enterPrintMode(url) {
    const mail = findMessage();
    const previousTitle = document.title;
    const wrapper = document.createElement("main");
    wrapper.id = "olm-pdf-document";
    const link = document.createElement("a");
    link.className = "olm-source-link";
    link.href = url;
    link.textContent = `元のメールを Outlook で開く: ${url}`;
    const heading = document.createElement("h1");
    heading.textContent = mail.subject;
    const content = mail.container.cloneNode(true);
    wrapper.append(link, heading, content);
    const style = document.createElement("style");
    style.id = "olm-pdf-style";
    style.textContent = `
      @page { size: A4; margin: 14mm; }
      body > * { display: none !important; }
      body > #olm-pdf-document { display: block !important; color: #1b1b1b; font: 12px/1.55 Arial, "Yu Gothic", sans-serif; }
      #olm-pdf-document * { max-width: 100% !important; }
      #olm-pdf-document h1 { font-size: 18px; border-bottom: 1px solid #bbb; padding-bottom: 8px; }
      #olm-pdf-document .olm-source-link { display: block; padding: 10px; margin-bottom: 14px; border: 1px solid #3574c5; background: #eef6ff; color: #0645ad; overflow-wrap: anywhere; }
      #olm-pdf-document button, #olm-pdf-document [role="button"] { display: none !important; }
    `;
    document.head.appendChild(style);
    document.body.appendChild(wrapper);
    document.title = mail.subject;
    cleanup = () => {
      wrapper.remove(); style.remove(); document.title = previousTitle; cleanup = null;
    };
    return { sender: mail.sender, subject: mail.subject, date: mail.date };
  }

  chrome.runtime.onMessage.addListener((request, _sender, respond) => {
    if (request.type === MESSAGE && request.action === "start") {
      try { respond({ ok: true, metadata: enterPrintMode(request.url) }); }
      catch (error) { respond({ ok: false, error: error.message }); }
      return;
    }
    if (request.type === MESSAGE && request.action === "end") {
      cleanup?.(); respond({ ok: true });
    }
  });
})();
