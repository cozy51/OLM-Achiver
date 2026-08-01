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
      if (found && elementText(found)) return found;
    }
    return null;
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
  const senderElement = first([
    '[data-testid="message-sender"]', '[data-testid="sender-name"]',
    '[aria-label^="差出人"]', '[aria-label^="From"]'
  ], container);
  const timeElement = first(['time[datetime]', '[data-testid="message-date"]', '[title*="202"]'], container);

  document.getElementById(WRAPPER_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
  globalThis.__olmPreviousTitle = document.title;
  const wrapper = document.createElement("main");
  wrapper.id = WRAPPER_ID;
  const link = document.createElement("a");
  link.className = "olm-source-link";
  link.href = sourceUrl;
  link.textContent = `元のメールを Outlook で開く: ${sourceUrl}`;
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
      sender: elementText(senderElement).replace(/^(差出人|From)\s*[:：]?\s*/i, "") || "差出人不明",
      date: timeElement?.getAttribute("datetime") || timeElement?.getAttribute("title") || elementText(timeElement) || new Date().toISOString()
    }
  };
}

export function buildPrintPageExpression(action, sourceUrl = "") {
  return `(${outlookPrintMode.toString()})(${JSON.stringify(action)}, ${JSON.stringify(sourceUrl)})`;
}
