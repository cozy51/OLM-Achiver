// Outlook の差出人まわりの DOM 構造を調べるための調査用スクリプトです。
// 拡張機能の一部ではありません。Outlook のタブで DevTools のコンソールを開き、
// SENDER に画面上の差出人名を入れてから、このファイル全体を貼り付けて実行します。
//
// 出力は構造(タグ・属性・座標)だけで、氏名・メールアドレス・件名・本文は
// <SENDER> や <MASKED> に置き換わります。属性名と、差出人を特定するために
// 必要な既定の語(差出人・宛先・プロフィールなど)は残ります。
(() => {
  // ↓ Outlook の画面に表示されている差出人名をそのまま貼り付けてください。
  const SENDER = "";
  const MASK = true;

  const KEEP = /^(?:差出人|送信者|宛先|件名|返信|全員に返信|転送|削除|その他|アプリ|リアクション|プロフィール|カード|開く|添付ファイル|未読|開封済み|フラグ|重要|下書き|メッセージ|本文|閲覧ウィンドウ|From|Sender|To|Cc|Bcc|Subject|Reply|Forward|Delete|More|Apps|Profile|card|Open|Attachment|Unread|Flag|Draft|Message|body|Reading|pane)$/i;

  const mask = value => {
    let text = String(value == null ? "" : value);
    if (!MASK) return text.slice(0, 160);
    if (SENDER) text = text.split(SENDER).join("<SENDER>");
    text = text.replace(/[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+/g, "<MAIL>");
    // Mask word by word so the shape of a label such as "差出人: 山田 太郎"
    // survives while the person does not.
    text = text.replace(/[\p{L}\p{N}][\p{L}\p{N}'-]*/gu, word =>
      KEEP.test(word) || word === "SENDER" || word === "MAIL" ? word : "*".repeat(Math.min(word.length, 6)));
    return text.slice(0, 160);
  };

  const describe = element => {
    if (!element || element.nodeType !== 1) return "(none)";
    const parts = [element.tagName.toLowerCase()];
    if (element.id) parts.push(`#${/^[A-Za-z][\w-]{0,20}$/.test(element.id) ? element.id : "<generated>"}`);
    for (const attribute of element.attributes) {
      const name = attribute.name;
      // Generated class names change on every Outlook release and are noise.
      if (name === "class" || name === "id" || name === "style") continue;
      const value = /^(?:aria-label|title|alt|placeholder|value|href)$/i.test(name)
        ? mask(attribute.value) : attribute.value.slice(0, 60);
      parts.push(`[${name}="${value}"]`);
    }
    const rect = element.getBoundingClientRect();
    parts.push(`@${Math.round(rect.left)},${Math.round(rect.top)} ${Math.round(rect.width)}x${Math.round(rect.height)}`);
    const own = [...element.childNodes].filter(node => node.nodeType === 3).map(node => node.textContent.trim()).join(" ").trim();
    if (own) parts.push(`text="${mask(own)}"`);
    return parts.join(" ");
  };

  const lines = [];
  const say = (...values) => lines.push(values.join(" "));

  say("== URL ==", location.origin + location.pathname);

  say("", "== 拡張機能が探しているルート ==");
  for (const selector of [
    '[data-app-section="MailReadCompose"]',
    '[aria-label*="閲覧ウィンドウ"]',
    '[aria-label*="Reading pane"]',
    '[data-testid="message-body"]',
    '[aria-label="メッセージ本文"]',
    '[aria-label="Message body"]',
    '[role="document"]'
  ]) {
    say(`${selector} -> ${document.querySelector(selector) ? describe(document.querySelector(selector)) : "見つかりません"}`);
  }

  say("", "== 画面にある data-app-section / data-testid / automation-id ==");
  const kinds = new Map();
  for (const element of document.querySelectorAll("[data-app-section], [data-testid], [automation-id], [autoid]")) {
    for (const name of ["data-app-section", "data-testid", "automation-id", "autoid"]) {
      const value = element.getAttribute(name);
      if (value) kinds.set(`${name}="${value}"`, (kinds.get(`${name}="${value}"`) || 0) + 1);
    }
  }
  say([...kinds.entries()].map(([key, count]) => `${key} x${count}`).join("\n") || "(なし)");

  say("", "== 「宛先:」らしきラベル ==");
  let found = 0;
  for (const element of document.querySelectorAll("span, div, [aria-label]")) {
    const label = element.getAttribute("aria-label") || element.textContent.trim();
    if (!/^\s*(?:宛先|To|Cc)\s*[:：]/i.test(label) || label.length > 40) continue;
    if (found++ > 4) break;
    say(describe(element));
  }
  if (!found) say("(見つかりません)");

  say("", "== 差出人名を含む要素とその祖先 ==");
  if (!SENDER) {
    say("SENDER が空です。ファイル先頭の SENDER に差出人名を入れて実行してください。");
  } else {
    let hits = 0;
    const all = document.querySelectorAll("*");
    for (const element of all) {
      const own = [...element.childNodes].filter(node => node.nodeType === 3).map(node => node.textContent).join("");
      const attrs = ["aria-label", "title", "alt"].map(name => element.getAttribute(name) || "").join(" ");
      if (!own.includes(SENDER) && !attrs.includes(SENDER)) continue;
      if (hits++ > 5) break;
      say(`--- ${hits} ---`);
      let node = element;
      for (let depth = 0; node && depth < 8; depth += 1) {
        say(`${"  ".repeat(depth)}${describe(node)}`);
        node = node.parentElement;
      }
    }
    if (!hits) say("差出人名を含む要素が見つかりません。表示名と完全に同じ文字列か確認してください。");
  }

  const report = lines.join("\n");
  console.log(report);
  try { copy(report); console.log("(クリップボードにコピーしました)"); } catch { /* copy() is DevTools only */ }
  return `${lines.length} 行を出力しました。`;
})();
