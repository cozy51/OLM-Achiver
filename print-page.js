// attachments is the list read by scan-attachments.js and named by
// attachments.js: [{ name, url }], where an empty url means the file has no
// address of its own and is only reachable through the message.
export function outlookPrintMode(action, sourceUrl, attachments) {
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

  const toCloudDeepLink = value => {
    try {
      const url = new URL(value, location.origin);
      if (/\/mail\/deeplink\/read\//i.test(url.pathname)) {
        url.protocol = "https:";
        url.host = "outlook.cloud.microsoft";
        return url.href;
      }
      const idMatch = url.pathname.match(/\/mail\/(?:[^/]+\/)?id\/([^/]+)/i);
      if (!idMatch) return "";
      const itemId = decodeURIComponent(idMatch[1]);
      const pathId = itemId.replace(/\//g, "-").replace(/\+/g, "_").replace(/=+$/g, "");
      return `https://outlook.cloud.microsoft/mail/deeplink/read/${pathId}?ItemID=${encodeURIComponent(itemId)}&exvsurl=1`;
    } catch {
      return "";
    }
  };

  const findMessageUrl = currentUrl => {
    if (isMessageUrl(currentUrl)) return toCloudDeepLink(currentUrl);

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
        if (isMessageUrl(href)) return toCloudDeepLink(href);
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
      const pathId = decodedId.replace(/\//g, "-").replace(/\+/g, "_").replace(/=+$/g, "");
      return `https://outlook.cloud.microsoft/mail/deeplink/read/${pathId}?ItemID=${encodeURIComponent(decodedId)}&exvsurl=1`;
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

  // Removed before the search so the copy of the message inside a wrapper left
  // behind by an interrupted run cannot be mistaken for the live message.
  document.getElementById(WRAPPER_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();

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

  // The tiles in the body show the file names truncated and carry no URL, so
  // the names are listed once more as text that survives in the PDF. A file
  // stored in the message has no address of its own, and the link at the top of
  // the page already opens that message, so those rows stay plain names.
  const files = (Array.isArray(attachments) ? attachments : [])
    .map(item => ({ name: String(item?.name || "").trim(), url: item?.url || "", mail: item?.mail === true }))
    .filter(item => item.name);

  // The kinds worth telling apart at a glance, in Office's own colours so the
  // list reads the way the attachment well does. Anything else keeps its own
  // extension as the label, which says more than a generic icon would.
  const FILE_KINDS = [
    { match: /\.(?:xlsx|xlsm|xlsb|xls|csv)$/i, label: "XLS", color: "#107C41" },
    { match: /\.(?:docx|docm|doc|rtf)$/i, label: "DOC", color: "#185ABD" },
    { match: /\.(?:pptx|pptm|ppt|pps)$/i, label: "PPT", color: "#C43E1C" },
    { match: /\.pdf$/i, label: "PDF", color: "#C4314B" },
    { match: /\.(?:png|jpe?g|gif|bmp|tiff?|webp|heic|svg)$/i, label: "IMG", color: "#8764B8" },
    { match: /\.(?:zip|7z|rar|lzh|tar|gz)$/i, label: "ZIP", color: "#8A7B4F" },
    { match: /\.(?:txt|log|md|json|xml|html?)$/i, label: "TXT", color: "#5D5A58" },
    { match: /\.(?:dwg|dxf|step|stp|igs|iges|sldprt|sldasm|catpart)$/i, label: "CAD", color: "#00707F" },
    { match: /\.(?:msg|eml)$/i, mail: true }
  ];
  const MAIL_COLOR = "#0F6CBD";
  const OTHER_COLOR = "#6E6E6E";
  // Drawn as shapes rather than as a background colour: a background is what a
  // printer drops first, and the PDF is made through the print path. The nodes
  // are built one at a time because a build that enforces Trusted Types refuses
  // markup assigned through innerHTML.
  const SVG_NS = "http://www.w3.org/2000/svg";
  const svgNode = (name, attributes) => {
    const node = document.createElementNS(SVG_NS, name);
    for (const key of Object.keys(attributes)) node.setAttribute(key, String(attributes[key]));
    return node;
  };
  const attachmentIcon = item => {
    const kind = FILE_KINDS.find(entry => entry.match.test(item.name));
    // A mail attached to a mail has a subject where a file has a name, so what
    // it is can only come from the scan; ".msg" says the same thing in writing.
    // An extension outranks it: a row named like a file is drawn as that file.
    const isMail = Boolean(kind?.mail) || (item.mail && !kind);
    const icon = svgNode("svg", {
      class: "olm-attachment-icon", width: 16, height: 16, viewBox: "0 0 16 16", "aria-hidden": "true"
    });
    icon.append(svgNode("rect", {
      x: 0.5, y: 0.5, width: 15, height: 15, rx: 3, fill: isMail ? MAIL_COLOR : (kind?.color || OTHER_COLOR)
    }));
    if (isMail) {
      icon.append(svgNode("rect", {
        x: 3, y: 4.5, width: 10, height: 7.5, rx: 1, fill: "none", stroke: "#fff", "stroke-width": 1.2
      }));
      icon.append(svgNode("path", {
        d: "M3.4 5.1 L8 8.7 L12.6 5.1", fill: "none", stroke: "#fff",
        "stroke-width": 1.2, "stroke-linecap": "round", "stroke-linejoin": "round"
      }));
      return icon;
    }
    const extension = (item.name.match(/\.([A-Za-z0-9]{1,8})$/) || [])[1] || "";
    const label = kind?.label || (extension ? extension.slice(0, 4).toUpperCase() : "FILE");
    const text = svgNode("text", {
      x: 8, y: 11, "text-anchor": "middle", "font-family": "Arial, sans-serif",
      "font-size": label.length > 3 ? 5.2 : 6.6, "font-weight": 700, fill: "#fff"
    });
    text.textContent = label;
    icon.append(text);
    return icon;
  };

  globalThis.__olmPreviousTitle = document.title;
  const wrapper = document.createElement("main");
  wrapper.id = WRAPPER_ID;
  const link = document.createElement("a");
  link.className = "olm-source-link";
  link.href = messageUrl;
  link.textContent = `元のメールを Outlook で開く: ${messageUrl}`;
  const heading = document.createElement("h1");
  heading.textContent = subject;
  wrapper.append(link, heading);
  if (files.length) {
    const section = document.createElement("section");
    section.className = "olm-attachments";
    const title = document.createElement("h2");
    title.textContent = `添付ファイル (${files.length}件)`;
    const list = document.createElement("ol");
    for (const item of files) {
      const row = document.createElement("li");
      row.append(attachmentIcon(item));
      if (item.url) {
        const anchor = document.createElement("a");
        anchor.href = item.url;
        anchor.textContent = item.name;
        const address = document.createElement("span");
        address.className = "olm-attachment-url";
        // Printed on paper a link is lost, so the destination is written out.
        address.textContent = item.url;
        row.append(anchor, address);
      } else {
        const name = document.createElement("span");
        name.textContent = item.name;
        row.append(name);
      }
      list.append(row);
    }
    section.append(title, list);
    wrapper.append(section);
  }
  wrapper.append(container.cloneNode(true));
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = `
    @page { size: A4; margin: 14mm; }
    body > * { display: none !important; }
    body > #${WRAPPER_ID} { display: block !important; color: #1b1b1b; font: 12px/1.55 Arial, "Yu Gothic", sans-serif; }
    #${WRAPPER_ID} * { max-width: 100% !important; }
    #${WRAPPER_ID} h1 { font-size: 18px; border-bottom: 1px solid #bbb; padding-bottom: 8px; }
    #${WRAPPER_ID} .olm-source-link { display: block; padding: 10px; margin-bottom: 14px; border: 1px solid #3574c5; background: #eef6ff; color: #0645ad; overflow-wrap: anywhere; }
    #${WRAPPER_ID} .olm-attachments { margin: 0 0 14px; padding: 8px 12px; border: 1px solid #cfd8e3; background: #f6f8fb; }
    #${WRAPPER_ID} .olm-attachments h2 { font-size: 13px; margin: 0 0 6px; }
    #${WRAPPER_ID} .olm-attachments ol { margin: 0; padding-left: 20px; }
    #${WRAPPER_ID} .olm-attachments li { margin-bottom: 5px; overflow-wrap: anywhere; }
    #${WRAPPER_ID} .olm-attachment-icon { width: 16px; height: 16px; margin-right: 6px; vertical-align: -3px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
    #${WRAPPER_ID} .olm-attachment-url { display: block; color: #555; font-size: 10px; margin-left: 22px; }
    #${WRAPPER_ID} button, #${WRAPPER_ID} [role="button"] { display: none !important; }
  `;
  document.head.appendChild(style);
  document.body.appendChild(wrapper);
  document.title = subject;
  return {
    ok: true,
    metadata: {
      subject,
      date: receivedDate,
      attachments: files
    }
  };
}

export function buildPrintPageExpression(action, sourceUrl = "", attachments = []) {
  return `(${outlookPrintMode.toString()})(${JSON.stringify(action)}, ${JSON.stringify(sourceUrl)}, ${JSON.stringify(attachments)})`;
}
