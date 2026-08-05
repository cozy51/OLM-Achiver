// 添付ファイルの「誤検知」を調べるための調査用スクリプトです。拡張機能の一部で
// はありません。誤検知したメールを Outlook で開き、DevTools のコンソールにこの
// ファイル全体を貼り付けて実行し、出力をそのまま共有してください。
//
// 実行する画面が結果を左右します。拡張機能は「PDFを保存したその瞬間の画面」を
// 読むので、**PDFを保存したときと同じ画面**(通常のOutlook画面で、そのメールを
// 閲覧ウィンドウに開いた状態)で実行してください。PDF内のリンクから開いた単独
// 表示の画面は作りが違い、添付欄が0か所になることがあります。
//
// PDFに出た余分な名前が分かっているときは、下の SUSPECT にその文字列を入れて
// ください。添付欄が見つからない画面でも、その文字列がDOMのどこにあるかを追跡
// します。
//
// tools/inspect-attachments.js が「添付欄がどんなDOMか」を見るのに対し、こちらは
// 拡張機能と同じ収集・判定をその場で走らせ、検出した名前ごとに
//   ・どの要素の、どの属性(ラベル)から取ったか
//   ・どの判定経路(拡張子/件名/コマンド直前/ラベルそのまま)で採用されたか
//   ・添付欄からその要素までの道すじ
// を出します。余分な1件がどこから来たのかを、これで特定できます。
//
// 検出した「名前」はPDFに載る文字列そのものなので伏せません。それ以外の本文・
// 氏名・メールアドレスは伏せ字にします(MASK を false にすると伏せません)。
(() => {
  const MASK = true;
  // 例: const SUSPECT = "yasuhiro.nishi";
  const SUSPECT = "";
  const KEEP = /^(?:添付ファイル|画像添付ファイル|添付|ファイル|個|件|ダウンロード|すべて|保存|プレビュー|開く|共有|削除|その他|メニュー|オプション|画像|写真|文書|リンク|クラウド|OneDrive|SharePoint|Attachment|Attachments|File|Files|Download|Save|Open|Preview|Share|Delete|More|Image|Photo|Document|Link|Cloud|Outlook|アイテム|item|メール|mail|message)$/i;

  const mask = value => {
    const text = String(value == null ? "" : value).replace(/\s+/g, " ").trim();
    if (!MASK) return text.slice(0, 200);
    return text
      .replace(/[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+/g, "<MAIL>")
      .replace(/([\p{L}\p{N}][\p{L}\p{N} _'-]*)(\.[A-Za-z0-9]{1,8})(?![\p{L}\p{N}])/gu,
        (_all, stem, extension) => `${"*".repeat(Math.min(stem.trim().length, 8))}${extension}`)
      .replace(/[\p{L}\p{N}][\p{L}\p{N}'-]*/gu, word =>
        KEEP.test(word) || word === "MAIL" ? word : "*".repeat(Math.min(word.length, 6)))
      .slice(0, 200);
  };

  const tag = element => {
    if (!element || element.nodeType !== 1) return "(none)";
    const parts = [element.tagName.toLowerCase()];
    const role = element.getAttribute("role");
    if (role) parts.push(`role="${role}"`);
    const testid = element.getAttribute("data-testid") || element.getAttribute("automation-id");
    if (testid) parts.push(`testid="${testid}"`);
    if (element.id) parts.push(`#${element.id.replace(/[A-Za-z0-9+/=]{12,}/g, "<itemid>")}`);
    const rect = element.getBoundingClientRect();
    parts.push(`@${Math.round(rect.left)},${Math.round(rect.top)} ${Math.round(rect.width)}x${Math.round(rect.height)}`);
    return parts.join(" ");
  };

  const pathTo = (element, stop) => {
    const steps = [];
    for (let node = element; node && node !== stop && steps.length < 8; node = node.parentElement) {
      steps.unshift(tag(node));
    }
    return steps.join(" > ");
  };

  // ---- ここから下は拡張機能の写しです(scan-attachments.js / attachments.js) ----
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
  const NAMED_WELL = [
    '[id$="_ATTACHMENTS"]',
    '[data-testid*="attachment" i]',
    '[automation-id*="attachment" i]'
  ];
  const LABELLED_WELL = [
    '[aria-label*="添付ファイル"]',
    '[aria-label*="Attachment" i]'
  ];
  const WELL_LIST = [...NAMED_WELL, ...LABELLED_WELL];
  const findWells = (scope, selectors) => {
    const matched = [...scope.querySelectorAll(selectors.join(", "))];
    return matched.filter(well => !matched.some(other => other !== well && other.contains(well)));
  };
  const inPane = readingPane ? findWells(readingPane, WELL_LIST) : [];
  const outside = findWells(document, NAMED_WELL).filter(well => !well.closest('[role="option"]'));
  const wells = inPane.length ? inPane : outside;
  const usedFallback = !inPane.length && wells.length > 0;

  const NAMED = ["aria-label", "title", "alt", "download", "name"];
  const labelledBy = element => (element.getAttribute("aria-labelledby") || "")
    .split(/\s+/)
    .filter(Boolean)
    .map(id => document.getElementById(id)?.textContent?.trim() || "");

  const tiles = new Map();
  const tileOf = element => {
    const found = element.closest('[role="option"], [role="listitem"]') ||
      element.closest('a[href], [draggable="true"]');
    const tile = found || element;
    if (!tiles.has(tile)) tiles.set(tile, tiles.size);
    return { tile, index: tiles.get(tile), inTile: Boolean(found) };
  };

  // 拡張機能は文字列だけを持ち帰りますが、調査では「どの要素のどの属性か」も
  // 覚えておきます。誤検知の出どころはそこにしか書かれていないためです。
  const entries = [];
  for (const well of wells) {
    for (const element of [well, ...well.querySelectorAll("*")]) {
      if (entries.length >= 600) break;
      const sourced = [
        ...NAMED.map(name => [name, element.getAttribute(name)]),
        ...labelledBy(element).map(text => ["aria-labelledby", text]),
        ["(テキスト)", element.children.length ? "" : (element.textContent || "").trim()],
        ...[...element.attributes].filter(item => item.name.startsWith("data-")).map(item => [item.name, item.value])
      ].filter(([, label]) => label && label.length <= 200);
      if (!sourced.length) continue;
      const { tile, index, inTile } = tileOf(element);
      const anchor = tile.matches("a[href]") ? tile : tile.querySelector("a[href]");
      entries.push({ tile: index, tileElement: tile, element, well, own: element === tile, inTile, sourced, href: anchor?.href || "" });
    }
  }

  const EXTENSION = /\.[A-Za-z0-9]{1,8}(?![\p{L}\p{N}])/gu;
  const HAS_LETTER = /[A-Za-z]/;
  const FILE_SIZE = /^[\d.,]+\s*(?:B|KB|MB|GB|TB|バイト)$/i;
  const FIELDS = /[、]\s*/;
  const PREFIX = /^(?:(?:画像)?添付ファイル|Attachment)\s*[:：]?\s*|^(?:開く|プレビュー|ダウンロード|選択)\s+/i;
  const ITEM_KIND = /Outlook\s*(?:アイテム|item)|メール\s*メッセージ|mail\s*message/i;
  const TILE_COMMAND = /開く|ダウンロード|保存|プレビュー|削除|選択|その他|オプション|メニュー|^(?:Open|Download|Save|Preview|Delete|Select|More|Options|Menu)\b/i;
  const COMMAND_TOKEN = /\s(?:開く|プレビュー|ダウンロード|保存|削除|選択|その他|オプション|メニュー|Open|Preview|Download|Save|Delete|Select|More|Options|Menu)(?=\s|$)/;
  const ADDRESS = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+/g;
  const ADDRESS_ONLY = /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+$/;
  const clean = value => String(value || "").replace(/[\uE000-\uF8FF]/g, " ").replace(/\s+/g, " ").trim();
  const usableName = name => Boolean(name) && name.length <= 120 && !FILE_SIZE.test(name) && !/[\\/]/.test(name) &&
    !ADDRESS_ONLY.test(name);

  const nameFrom = value => {
    const text = clean(value);
    for (const field of text.split(FIELDS)) {
      const addresses = [...field.matchAll(ADDRESS)].map(match => [match.index, match.index + match[0].length]);
      for (const match of field.matchAll(EXTENSION)) {
        if (!HAS_LETTER.test(match[0])) continue;
        const end = match.index + match[0].length;
        if (addresses.some(([from, to]) => match.index < to && end > from)) continue;
        const start = addresses.filter(([, to]) => to <= match.index).map(([, to]) => to).pop() || 0;
        const name = field.slice(start, end).replace(PREFIX, "").trim();
        if (!usableName(name)) continue;
        if (name.length <= match[0].length) continue;
        return name;
      }
    }
    return "";
  };
  const subjectFrom = value => {
    const text = clean(value);
    const kind = text.match(ITEM_KIND);
    if (!kind) return "";
    const subject = text.slice(0, kind.index).replace(/[、,]\s*$/, "").trim();
    return usableName(subject) ? subject : "";
  };
  const nameBeforeCommand = value => {
    const text = clean(value);
    const command = text.match(COMMAND_TOKEN);
    if (!command) return "";
    const name = text.slice(0, command.index).trim();
    return usableName(name) && /[\p{L}\p{N}]/u.test(name) ? name : "";
  };
  const looseName = value => {
    const text = clean(value);
    if (!usableName(text) || ITEM_KIND.test(text) || TILE_COMMAND.test(text)) return "";
    return /[\p{L}\p{N}]/u.test(text) ? text : "";
  };
  // ---- 写しはここまで ----

  // 拡張機能の nameOfTile と同じ順番で試し、当たった経路と元ラベルを残します。
  const ROUTES = [
    { label: "拡張子(attachmentNameFrom)", pick: nameFrom, ownOnly: false, needsTile: false },
    { label: "Outlookアイテムの件名", pick: subjectFrom, ownOnly: false, needsTile: false },
    { label: "コマンド直前", pick: nameBeforeCommand, ownOnly: true, needsTile: true },
    { label: "ラベルそのまま(タイル自身)", pick: looseName, ownOnly: true, needsTile: true },
    { label: "ラベルそのまま(タイル内すべて)", pick: looseName, ownOnly: false, needsTile: true }
  ];
  const decide = group => {
    const anyTile = group.some(entry => entry.inTile);
    for (const route of ROUTES) {
      if (route.needsTile && !anyTile) return { name: "", route: "(タイルの裏づけなし)" };
      for (const entry of group) {
        if (route.ownOnly && !entry.own) continue;
        for (const [attribute, value] of entry.sourced) {
          const name = route.pick(value);
          if (name) return { name, route: route.label, attribute, value, entry };
        }
      }
    }
    return { name: "", route: "(名前なし)" };
  };

  const grouped = new Map();
  for (const entry of entries) {
    if (!grouped.has(entry.tile)) grouped.set(entry.tile, []);
    grouped.get(entry.tile).push(entry);
  }

  const lines = [];
  const say = (...values) => lines.push(values.join(" "));
  say("== URL ==", location.origin + location.pathname);
  say("== 閲覧ウィンドウ ==", readingPane ? tag(readingPane) : "(見つかりません。ページ全体を対象にしました)");
  say("== 添付欄 ==", `${wells.length}か所`, usedFallback ? "(閲覧ウィンドウ内に無く、ページ全体から拾いました)" : "");
  // どのセレクターが効いていないのかは、合計だけでは分かりません。
  for (const selector of WELL_LIST) {
    const outsideCount = [...document.querySelectorAll(selector)].filter(element => !readingPane?.contains(element)).length;
    const kind = NAMED_WELL.includes(selector) ? "" : " ※閲覧ウィンドウ内でのみ採用(メール一覧の行も同じ語を持つため)";
    say(`  ${selector} -> 閲覧ウィンドウ内 ${root.querySelectorAll(selector).length}件 / 外 ${outsideCount}件${kind}`);
  }
  wells.forEach((well, index) => say(`  well[${index}]`, tag(well), `子孫${well.querySelectorAll("*").length}個`));
  if (!wells.length) {
    say("  ※ この画面では拡張機能も0件になります。PDFを保存したときと同じ画面で実行してください。");
  }

  say("", "== 検出結果(PDFに載る一覧と同じ) ==");
  let listed = 0;
  for (const [index, group] of grouped) {
    const result = decide(group);
    if (!result.name) continue;
    listed += 1;
    say(`[${listed}] 「${result.name}」`);
    say(`     経路: ${result.route}`);
    say(`     元ラベル: ${result.attribute}="${mask(result.value)}"`);
    say(`     その要素: ${tag(result.entry.element)}`);
    say(`     タイル: ${tag(result.entry.tileElement)} inTile=${result.entry.inTile}`);
    say(`     添付欄からの道すじ: ${pathTo(result.entry.element, result.entry.well.parentElement)}`);
    const url = entries.filter(entry => entry.tile === index).map(entry => entry.href).find(Boolean) || "";
    say(`     リンク: ${url ? `${url.slice(0, 40)}…(${url.length}文字)` : "(なし)"}`);
  }
  if (!listed) say("(1件も検出されませんでした)");

  say("", "== 名前にならなかったタイル ==");
  let skipped = 0;
  for (const group of grouped.values()) {
    const result = decide(group);
    if (result.name || skipped++ > 15) continue;
    say(`  ${result.route} <- ${tag(group[0].element)} ラベル例: "${mask(group[0].sourced[0]?.[1])}"`);
  }
  if (!skipped) say("  (なし)");

  // 添付欄が見つからない画面でも、余分な名前の出どころは追えます。文字列がどこに
  // 書かれているか(タイルのラベルか、引用された本文か)が分かれば足ります。
  say("", "== 手がかりの追跡 ==");
  if (!SUSPECT) {
    say('  (先頭の SUSPECT にPDFへ出た余分な名前を入れると、その文字列の在りかを追います)');
  } else {
    let found = 0;
    for (const element of document.querySelectorAll("*")) {
      const where = [];
      for (const attribute of element.attributes) {
        if (/^(?:class|style|href|src)$/i.test(attribute.name)) continue;
        if (attribute.value.includes(SUSPECT)) where.push(attribute.name);
      }
      const own = [...element.childNodes].filter(node => node.nodeType === 3).map(node => node.textContent).join(" ");
      if (own.includes(SUSPECT)) where.push("(テキスト)");
      if (!where.length || found++ >= 10) continue;
      say(`  ${where.join(" ")} <- ${tag(element)}`);
      say(`     文字列: "${mask(own.includes(SUSPECT) ? own : element.getAttribute(where[0]))}"`);
      say(`     道すじ: ${pathTo(element, document.body)}`);
      say(`     添付欄の中か: ${wells.some(well => well.contains(element)) ? "はい" : "いいえ"}`);
      const inBody = element.closest('[data-testid="message-body"], [aria-label="メッセージ本文"], [aria-label="Message body"], [role="document"]');
      say(`     メール本文の中か: ${inBody ? "はい(引用された本文の可能性)" : "いいえ"}`);
    }
    if (!found) say(`  「${SUSPECT}」はこの画面のDOMに見当たりません。保存時と同じ画面で実行してください。`);
  }

  const report = lines.join("\n");
  console.log(report);
  try { copy(report); console.log("(クリップボードにコピーしました)"); } catch { /* copy() は DevTools 専用 */ }
  return `${lines.length} 行を出力しました。`;
})();
