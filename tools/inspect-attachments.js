// Outlook の添付ファイルまわりの DOM 構造を調べるための調査用スクリプトです。
// 拡張機能の一部ではありません。添付ファイル付きのメールを開き、DevTools の
// コンソールにこのファイル全体を貼り付けて実行します。
//
// 出力は構造(タグ・属性・座標)だけです。ファイル名・氏名・本文は伏せますが、
// 拡張子(.jpg など)と文字数の形は残します。どこにファイル名が入っているかを
// 判断するために必要で、それだけでは中身が分からないためです。
(() => {
  const MASK = true;
  // 添付欄の構造を読むために必要な既定の語は残します。
  const KEEP = /^(?:添付ファイル|画像添付ファイル|添付|ファイル|個|件|ダウンロード|すべて|保存|プレビュー|開く|共有|削除|その他|メニュー|オプション|画像|写真|文書|リンク|クラウド|OneDrive|SharePoint|Attachment|Attachments|File|Files|Download|Save|Open|Preview|Share|Delete|More|Image|Photo|Document|Link|Cloud)$/i;

  const mask = value => {
    const text = String(value == null ? "" : value);
    if (!MASK) return text.slice(0, 160);
    return text
      .replace(/[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+/g, "<MAIL>")
      // A file name is masked to its shape: "見積書.xlsx" -> "***.xlsx". The
      // extension is what the extension keys on, so it has to stay readable.
      .replace(/([\p{L}\p{N}][\p{L}\p{N} _'-]*)(\.[A-Za-z0-9]{1,8})(?![\p{L}\p{N}])/gu,
        (_all, stem, extension) => `${"*".repeat(Math.min(stem.trim().length, 8))}${extension}`)
      .replace(/[\p{L}\p{N}][\p{L}\p{N}'-]*/gu, word =>
        KEEP.test(word) || word === "MAIL" ? word : "*".repeat(Math.min(word.length, 6)))
      .slice(0, 160);
  };

  const describe = element => {
    if (!element || element.nodeType !== 1) return "(none)";
    const parts = [element.tagName.toLowerCase()];
    if (element.id) parts.push(`#${element.id.replace(/[A-Za-z0-9+/=]{12,}/g, "<itemid>")}`);
    for (const attribute of element.attributes) {
      const name = attribute.name;
      // Generated class names change on every Outlook release and are noise.
      if (name === "class" || name === "id" || name === "style") continue;
      const raw = attribute.value;
      const value = /^(?:aria-label|title|alt|download|name|placeholder|value)$/i.test(name) ? mask(raw)
        : /^(?:href|src)$/i.test(name) ? `${raw.slice(0, 40)}…(${raw.length}文字)`
        : raw.slice(0, 60);
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

  say("", "== 拡張機能が添付欄を探すセレクター ==");
  const WELL = [
    '[id$="_ATTACHMENTS"]',
    '[data-testid*="attachment" i]',
    '[automation-id*="attachment" i]',
    '[aria-label*="添付ファイル"]',
    '[aria-label*="Attachment" i]'
  ];
  const wells = [];
  for (const selector of WELL) {
    const matched = [...document.querySelectorAll(selector)];
    say(`${selector} -> ${matched.length}件`);
    for (const element of matched) if (!wells.includes(element)) wells.push(element);
  }

  say("", "== 添付欄の中身(見つかった順に最大2か所・深さ4まで) ==");
  if (!wells.length) {
    say("(添付欄が1つも見つかりません。添付ファイル付きのメールを開いているか確認してください)");
  }
  const walk = (element, depth, limit) => {
    say(`${"  ".repeat(depth)}${describe(element)}`);
    if (depth >= limit) {
      if (element.children.length) say(`${"  ".repeat(depth + 1)}…(子要素 ${element.children.length}件は省略)`);
      return;
    }
    for (const child of [...element.children].slice(0, 12)) walk(child, depth + 1, limit);
  };
  for (const well of wells.slice(0, 2)) {
    say("---");
    walk(well, 0, 4);
  }

  say("", "== 添付欄の中でファイル名らしき文字列を持つ要素 ==");
  // The extension keeps only strings that end in an extension, so this is the
  // same question it asks: does the name reach the DOM at all, and where?
  const FILE_NAME = /[^\s、,]{1,80}\.[A-Za-z0-9]{1,8}(?![\p{L}\p{N}])/u;
  let hits = 0;
  for (const well of wells) {
    for (const element of [well, ...well.querySelectorAll("*")]) {
      const sources = [];
      for (const attribute of element.attributes) {
        if (/^(?:class|style|id|href|src)$/i.test(attribute.name)) continue;
        if (FILE_NAME.test(attribute.value)) sources.push(attribute.name);
      }
      const own = [...element.childNodes].filter(node => node.nodeType === 3).map(node => node.textContent).join(" ");
      if (FILE_NAME.test(own)) sources.push("(テキスト)");
      if (!sources.length || hits++ > 20) continue;
      say(`${sources.join(" ")} <- ${describe(element)}`);
    }
  }
  if (!hits) say("(ファイル名らしき文字列がどこにもありません。名前がDOMに出ていない可能性があります)");

  say("", "== 添付欄の中の role / aria-label 一覧 ==");
  const roles = new Map();
  for (const well of wells) {
    for (const element of [well, ...well.querySelectorAll("*")]) {
      const key = `${element.tagName.toLowerCase()} role="${element.getAttribute("role") || ""}" aria-label有無=${element.hasAttribute("aria-label")}`;
      roles.set(key, (roles.get(key) || 0) + 1);
    }
  }
  say([...roles.entries()].map(([key, count]) => `${key} x${count}`).join("\n") || "(なし)");

  const report = lines.join("\n");
  console.log(report);
  try { copy(report); console.log("(クリップボードにコピーしました)"); } catch { /* copy() is DevTools only */ }
  return `${lines.length} 行を出力しました。`;
})();
