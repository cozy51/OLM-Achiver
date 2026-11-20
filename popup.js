import { describeAttachments, normalizeSaveFormat, storedSaveFormat } from "./lib.js";

const button = document.querySelector("#save");
const status = document.querySelector("#status");
const senderInput = document.querySelector("#sender");
const pasteButton = document.querySelector("#paste");
const hint = document.querySelector("#hint");
const commentInput = document.querySelector("#comment");
const formatRadios = [...document.querySelectorAll('input[name="format"]')];
const FORMAT_KEY = "saveFormat";
const BUTTON_LABELS = {
  pdf: "📩 このメールを PDF 保存",
  markdown: "📩 このメールを Markdown 保存",
  both: "📩 このメールを PDF と Markdown で保存"
};

function selectedFormat() {
  return normalizeSaveFormat(formatRadios.find(radio => radio.checked)?.value);
}

function showFormat(format) {
  for (const radio of formatRadios) radio.checked = radio.value === format;
  button.textContent = BUTTON_LABELS[format];
}

// The choice is kept between uses: someone who archives in Markdown wants it
// every time, and should never have to pick it again.
showFormat("pdf");
chrome.storage.local.get(["saveFormat", "saveMarkdown"])
  .then(stored => showFormat(storedSaveFormat(stored)))
  .catch(() => {});
for (const radio of formatRadios) {
  radio.addEventListener("change", () => {
    const format = selectedFormat();
    showFormat(format);
    chrome.storage.local.set({ [FORMAT_KEY]: format }).catch(() => {});
  });
}

// The manifest is the single source of truth for the version, so the heading
// never drifts from what the browser actually installed.
document.querySelector("#version").textContent = `v${chrome.runtime.getManifest().version}`;

const MANUAL_HINT = "Outlookに表示されている差出人名をコピーし、ここへ貼り付けてください。";

// Only the best match is shown. Listing the runners-up alongside it invites
// the question of which one is right, when the field itself already answers it.
async function detectSender() {
  try {
    const result = await chrome.runtime.sendMessage({ type: "detect-sender" });
    if (!result?.ok) throw new Error(result?.error || "差出人を取得できませんでした。");
    if (!result.sender) {
      const reasons = [];
      if (!result.diagnostics?.readingPane) reasons.push("閲覧ウィンドウを特定できません");
      if (!result.diagnostics?.recipientRow) reasons.push("宛先行を特定できません");
      hint.textContent = reasons.length
        ? `差出人を検出できません(${reasons.join("・")})。${MANUAL_HINT}`
        : MANUAL_HINT;
      return;
    }
    if (senderInput.value) return;
    senderInput.value = result.sender.name;
    hint.textContent = result.confident
      ? "差出人を自動入力しました。違っていれば直接修正してください。"
      : "差出人の候補を入力しました。念のため確認してください。";
  } catch {
    hint.textContent = MANUAL_HINT;
  }
}

detectSender();

pasteButton.addEventListener("click", async () => {
  status.className = "";
  try {
    const clipboardText = (await navigator.clipboard.readText()).trim();
    if (!clipboardText) throw new Error("クリップボードに文字がありません。");
    senderInput.value = clipboardText.replace(/\s+/g, " ");
    status.textContent = "差出人名を貼り付けました。";
    senderInput.focus();
  } catch (error) {
    status.className = "error";
    status.textContent = `貼り付けできませんでした: ${error.message || "クリップボードを確認してください。"}`;
  }
});

button.addEventListener("click", async () => {
  const senderName = senderInput.value.trim();
  if (!senderName) {
    status.className = "error";
    status.textContent = "差出人名を候補から選ぶか、Outlookからコピーして貼り付けてください。";
    senderInput.focus();
    return;
  }
  button.disabled = true;
  status.className = "";
  const format = selectedFormat();
  status.textContent = format === "markdown" ? "Markdown を作成しています…" : "PDF を作成しています…";
  try {
    const result = await chrome.runtime.sendMessage({ type: "save-current-mail", senderName, comment: commentInput.value.trim(), format });
    if (!result.ok) throw new Error(result.error || "保存に失敗しました。");
    const saved = !result.filename ? ""
      : result.printDialog
        ? `Edgeの印刷画面を開きました。「プリンター」で「PDFとして保存」を選んでください。\n推奨ファイル名: ${result.filename}`
        : `保存しました: ${result.filename}`;
    const attachments = describeAttachments(result.attachments, { wells: result.attachmentWells, document: format === "markdown" ? "Markdown" : "PDF" });
    const markdown = result.markdownFilename
      ? `Markdownを「ダウンロード」フォルダーへ保存しました: ${result.markdownFilename}`
      : result.markdownError ? `Markdownは保存できませんでした: ${result.markdownError}` : "";
    status.textContent = [saved, attachments, markdown].filter(Boolean).join("\n");
  } catch (error) {
    status.className = "error";
    const message = error.message || "保存に失敗しました。";
    status.textContent = `${message}\n\n「失敗箇所」を確認すると、URL取得前かPDF作成時かを判別できます。`;
  }
  button.disabled = false;
});
