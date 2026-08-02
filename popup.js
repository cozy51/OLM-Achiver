const button = document.querySelector("#save");
const status = document.querySelector("#status");
const senderInput = document.querySelector("#sender");
const pasteButton = document.querySelector("#paste");
const hint = document.querySelector("#hint");

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
  status.textContent = "PDF を作成しています…";
  try {
    const result = await chrome.runtime.sendMessage({ type: "save-current-mail", senderName });
    if (!result.ok) throw new Error(result.error || "保存に失敗しました。");
    status.textContent = result.printDialog
      ? `Edgeの印刷画面を開きました。「プリンター」で「PDFとして保存」を選んでください。\n推奨ファイル名: ${result.filename}`
      : `保存しました: ${result.filename}`;
  } catch (error) {
    status.className = "error";
    const message = error.message || "保存に失敗しました。";
    status.textContent = `${message}\n\n「失敗箇所」を確認すると、URL取得前かPDF作成時かを判別できます。`;
  }
  button.disabled = false;
});
