const button = document.querySelector("#save");
const status = document.querySelector("#status");
const senderInput = document.querySelector("#sender");
const pasteButton = document.querySelector("#paste");

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
    status.textContent = "Outlookから差出人名をコピーして貼り付けてください。";
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
