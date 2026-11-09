import { parseMailDate } from "./lib.js";

// Runs inside the Outlook page through chrome.scripting, so it is serialised on
// its own: nothing from this module's scope may be referenced in its body.
//
// The aim is a file that reads well as plain text and renders tidily, not a
// copy of the mail's layout. Outlook writes one <div> (or <p class=MsoNormal>)
// per line and an empty one between paragraphs, so lines are kept as hard line
// breaks and the empty blocks become the paragraph gaps.
export function extractBodyMarkdown() {
  const WRAPPER_ID = "olm-pdf-document";
  const SKIP = new Set([
    "style", "script", "noscript", "template", "head", "title", "meta", "link",
    "button", "input", "select", "textarea", "svg", "img", "picture", "video",
    "audio", "canvas", "iframe", "object"
  ]);
  const BLOCK = new Set([
    "div", "section", "article", "header", "footer", "main", "aside", "nav",
    "address", "center", "form", "fieldset", "figure", "figcaption", "dl", "dt",
    "dd", "tbody", "thead", "tfoot", "tr", "td", "th", "caption", "li", "table"
  ]);
  const BODY_SELECTORS = [
    '[data-testid="message-body"]',
    '[aria-label="メッセージ本文"]',
    '[aria-label="Message body"]',
    '[role="document"]'
  ];
  const PANE_SELECTORS = [
    '[data-app-section="MailReadCompose"]',
    '[aria-label*="閲覧ウィンドウ"]',
    '[aria-label*="Reading pane"]'
  ];

  // The copy of the message placed for printing is never the message itself.
  const pick = (selectors, root) => {
    for (const selector of selectors) {
      for (const element of root.querySelectorAll(selector)) {
        if (!element.closest(`#${WRAPPER_ID}`) && element.textContent.trim()) return element;
      }
    }
    return null;
  };
  const pane = pick(PANE_SELECTORS, document);
  const body = pick(BODY_SELECTORS, pane || document);
  if (!body) return { ok: false, markdown: "" };

  const hidden = element => {
    try { return getComputedStyle(element).display === "none"; } catch { return false; }
  };
  // Full-width spaces are kept: Japanese mail indents with them on purpose.
  const collapse = value => value.replace(/[ \t\n\r\f ]+/g, " ");
  const escapeText = value => value
    .replace(/\\(?=[!-/:-@[-`{-~])/g, "\\\\")
    .replace(/([*`])/g, "\\$1");
  // A line of a mail that happens to start like Markdown syntax must stay text,
  // or a "- 項目" line would turn the rest of the paragraph into a list.
  const escapeLineStart = value => value
    .replace(/^(#{1,6})(?=\s|$)/, "\\$1")
    .replace(/^>/, "\\>")
    .replace(/^([-+])(?=\s)/, "\\$1")
    .replace(/^(\d+)([.)])(?=\s)/, "$1\\$2");
  const isRule = value => /^(?:[-=_]\s*){3,}$/.test(value) || /^(?:\\\*\s*){3,}$/.test(value);
  const isEmpty = element => !element.textContent.replace(/[\s 　]+/g, "") &&
    !element.querySelector("hr, table");
  const linkTarget = anchor => {
    let href = anchor.href || "";
    try {
      const url = new URL(href);
      // Outlook rewrites links through Safe Links; the original is what a
      // reader of the archive wants to see and follow.
      if (/(^|\.)safelinks\.protection\.outlook\.com$/i.test(url.hostname) && url.searchParams.get("url")) {
        href = url.searchParams.get("url");
      }
    } catch { return ""; }
    if (!/^(https?:|mailto:)/i.test(href)) return "";
    return href.replace(/[ ()<>]/g, character => encodeURIComponent(character));
  };
  // CommonMark leaves "**" unrendered next to punctuation in CJK text, so only
  // text that starts and ends with a letter or digit is marked as bold.
  const canEmphasise = value => /^[\p{L}\p{N}]/u.test(value) && /[\p{L}\p{N}]$/u.test(value) && !value.includes("**");

  const render = root => {
    const blocks = [];
    let lines = [];
    let line = "";
    const endLine = () => {
      const text = line.replace(/ {2,}/g, " ").replace(/^ +| +$/g, "");
      line = "";
      if (!text) return false;
      if (isRule(text)) {
        endParagraph();
        blocks.push("---");
      } else {
        lines.push(escapeLineStart(text));
      }
      return true;
    };
    const endParagraph = () => {
      if (line) endLine();
      if (lines.length) blocks.push(lines.join("  \n"));
      lines = [];
    };
    const solid = text => {
      endParagraph();
      if (text) blocks.push(text);
    };
    // Runs the children and hands back what they added to the current line,
    // or null when they broke the line, which leaves nothing safe to wrap.
    const inline = node => {
      const before = { blocks: blocks.length, lines: lines.length, at: line.length };
      walk(node);
      if (blocks.length !== before.blocks || lines.length !== before.lines || line.length < before.at) return null;
      const segment = line.slice(before.at);
      line = line.slice(0, before.at);
      return segment;
    };

    const visit = node => {
      if (node.nodeType === Node.TEXT_NODE) {
        line += escapeText(collapse(node.data));
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const tag = node.localName;
      if (SKIP.has(tag) || node.id === WRAPPER_ID || hidden(node)) return;

      if (/divRplyFwdMsg$/i.test(node.id || "")) {
        // The header of the quoted message: a rule marks where it begins.
        solid("---");
        walk(node);
        endParagraph();
        return;
      }
      switch (tag) {
        case "br":
          if (!endLine()) endParagraph();
          return;
        case "hr":
          solid("---");
          return;
        case "h1": case "h2": case "h3": case "h4": case "h5": case "h6": {
          const text = render(node).replace(/ {2}\n/g, " ").replace(/\n+/g, " ").trim();
          // The document's own title and sections sit above the body's headings.
          if (text) solid(`${"#".repeat(Math.min(6, Number(tag[1]) + 2))} ${text}`);
          return;
        }
        case "pre": {
          const text = node.textContent.replace(/\n+$/, "");
          if (text.trim()) solid(`\`\`\`\n${text}\n\`\`\``);
          return;
        }
        case "blockquote": {
          const text = render(node);
          if (text) solid(text.split("\n").map(row => row ? `> ${row}` : ">").join("\n"));
          return;
        }
        case "ul": case "ol": {
          const ordered = tag === "ol";
          const start = Number(node.getAttribute("start")) || 1;
          const items = [...node.children].filter(child => child.localName === "li" && !hidden(child));
          const text = items.map((item, index) => {
            const marker = ordered ? `${start + index}. ` : "- ";
            // An item stays one tight entry, nested list included.
            const content = (render(item) || "").replace(/\n{2,}/g, "\n");
            return marker + content.split("\n")
              .map((row, number) => number && row ? " ".repeat(marker.length) + row : row)
              .join("\n");
          }).join("\n");
          solid(text);
          return;
        }
        case "table":
          if (renderTable(node)) return;
          break;
        case "code": case "kbd": case "tt": case "samp": {
          const text = collapse(node.textContent).trim();
          if (text && !text.includes("`")) line += `\`${text}\``;
          else walk(node);
          return;
        }
        case "b": case "strong": {
          const segment = inline(node);
          if (segment === null) return;
          const [, lead, core, tail] = segment.match(/^( *)(.*?)( *)$/s);
          line += canEmphasise(core) ? `${lead}**${core}**${tail}` : segment;
          return;
        }
        case "a": {
          const href = linkTarget(node);
          const segment = inline(node);
          if (segment === null) return;
          const core = segment.trim();
          const shown = node.textContent.trim();
          if (!href || !core) {
            line += segment;
          } else if (/^mailto:/i.test(href) || shown === href || shown.replace(/\/$/, "") === href.replace(/\/$/, "")) {
            // A written-out address already says where it goes.
            line += /^mailto:/i.test(href) ? segment : `<${href}>`;
          } else {
            line += `[${core.replace(/([[\]])/g, "\\$1")}](${href})`;
          }
          return;
        }
        default:
          break;
      }

      // Outlook desktop writes every line as its own <p class=MsoNormal>, so
      // only other paragraphs are kept apart by a blank line.
      const paragraph = tag === "p" && !/\bMso/i.test(node.className || "");
      if (tag === "p" || BLOCK.has(tag)) {
        if (isEmpty(node)) {
          endParagraph();
          return;
        }
        if (paragraph) endParagraph(); else endLine();
        walk(node);
        if (paragraph) endParagraph(); else endLine();
        return;
      }
      walk(node);
    };
    const walk = node => {
      for (const child of node.childNodes) visit(child);
    };

    // Only a table that holds data becomes a Markdown table. Mail layouts and
    // signatures are built from tables too, and those read better as lines.
    const renderTable = table => {
      if (table.querySelector("table")) return false;
      const cell = element => render(element).replace(/ {2}\n/g, "<br>").replace(/\n+/g, "<br>").replace(/\|/g, "\\|").trim();
      let rows = [...table.rows]
        .filter(row => !hidden(row))
        .map(row => [...row.cells].map(cell));
      rows = rows.filter(row => row.some(Boolean));
      const width = Math.max(0, ...rows.map(row => row.length));
      const used = [...Array(width).keys()].filter(column => rows.some(row => row[column]));
      if (rows.length < 2 || used.length < 2) return false;
      if ([...table.querySelectorAll("td, th")].some(element => element.querySelector("ul, ol, blockquote, pre"))) return false;
      const lineOf = row => `| ${used.map(column => row[column] || " ").join(" | ")} |`;
      solid([
        lineOf(rows[0]),
        `| ${used.map(() => "---").join(" | ")} |`,
        ...rows.slice(1).map(lineOf)
      ].join("\n"));
      return true;
    };

    walk(root);
    endParagraph();
    return blocks.join("\n\n").replace(/\n{3,}/g, "\n\n").trim();
  };

  return { ok: true, markdown: render(body) };
}

const KIND_MARKS = [
  { match: /\.(?:xlsx|xlsm|xlsb|xls|csv)$/i, mark: "📊" },
  { match: /\.(?:docx|docm|doc|rtf)$/i, mark: "📝" },
  { match: /\.(?:pptx|pptm|ppt|pps)$/i, mark: "📽️" },
  { match: /\.pdf$/i, mark: "📕" },
  { match: /\.(?:png|jpe?g|gif|bmp|tiff?|webp|heic|svg)$/i, mark: "🖼️" },
  { match: /\.(?:zip|7z|rar|lzh|tar|gz)$/i, mark: "🗜️" },
  { match: /\.(?:txt|log|md|json|xml|html?)$/i, mark: "📄" },
  { match: /\.(?:dwg|dxf|step|stp|igs|iges|sldprt|sldasm|catpart)$/i, mark: "📐" },
  { match: /\.(?:msg|eml)$/i, mark: "✉️" }
];

// The same kinds the PDF marks with a coloured badge, as a character that
// survives in plain text. A mail attached to a mail has a subject for a name,
// so the scan's word for it decides when no extension does.
export function attachmentMark(item) {
  const name = String(item?.name || "");
  const kind = KIND_MARKS.find(entry => entry.match.test(name));
  if (kind) return kind.mark;
  return item?.mail ? "✉️" : "📎";
}

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

export function formatMailDate(value) {
  const instant = parseMailDate(value);
  if (!instant) return String(value || "").trim() || "不明";
  const pad = number => String(number).padStart(2, "0");
  return `${instant.getFullYear()}-${pad(instant.getMonth() + 1)}-${pad(instant.getDate())} (${WEEKDAYS[instant.getDay()]}) ${pad(instant.getHours())}:${pad(instant.getMinutes())}`;
}

const inlineText = value => String(value || "").replace(/\s+/g, " ").trim();
const escapeInline = value => inlineText(value).replace(/([\\*`[\]])/g, "\\$1");
const linkDestination = value => String(value || "").replace(/[ ()<>]/g, character => encodeURIComponent(character));

export function buildMarkdownDocument({ subject, sender, date, url, attachments, body }) {
  const title = inlineText(subject) || "件名なし";
  const files = (attachments || []).filter(item => inlineText(item?.name));
  const parts = [
    `# ${title.replace(/^#/, "\\#")}`,
    [
      `- **差出人**: ${escapeInline(sender) || "不明"}`,
      `- **受信日時**: ${formatMailDate(date)}`,
      url ? `- **元メール**: [Outlook で開く](${linkDestination(url)})` : null,
      files.length ? `- **添付**: ${files.length}件` : null
    ].filter(Boolean).join("\n")
  ];
  if (files.length) {
    parts.push(
      `## 📎 添付ファイル (${files.length}件)`,
      files.map((item, index) => {
        const name = escapeInline(item.name);
        const label = item.url ? `[${name}](${linkDestination(item.url)})` : name;
        return `${index + 1}. ${attachmentMark(item)} ${label}`;
      }).join("\n")
    );
  }
  parts.push("---", "## 本文", String(body || "").trim() || "*（本文を読み取れませんでした）*");
  return `${parts.join("\n\n")}\n`;
}

export function markdownFilename(pdfFilename) {
  return String(pdfFilename || "").replace(/\.pdf$/i, "") + ".md";
}

export function markdownDataUrl(text) {
  return `data:text/markdown;charset=utf-8,${encodeURIComponent(text)}`;
}
