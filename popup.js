const button = document.querySelector("#save");
const status = document.querySelector("#status");
button.addEventListener("click", async () => {
  button.disabled = true;
  status.className = "";
  status.textContent = "PDF を作成しています…";
  try {
    const result = await chrome.runtime.sendMessage({ type: "save-current-mail" });
    if (!result.ok) throw new Error(result.error || "保存に失敗しました。");
    status.textContent = `保存しました: ${result.filename}`;
  } catch (error) {
    status.className = "error";
    const message = error.message || "保存に失敗しました。";
    status.textContent = `${message}\n\n「失敗箇所」を確認すると、URL取得前かPDF作成時かを判別できます。`;
  }
  button.disabled = false;
});
