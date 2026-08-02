const button = document.querySelector("#save");
const status = document.querySelector("#status");
const senderInput = document.querySelector("#sender");
const pasteButton = document.querySelector("#paste");
const hint = document.querySelector("#hint");
const candidateBox = document.querySelector("#candidates");

const MANUAL_HINT = "Outlookに表示されている差出人名をコピーし、ここへ貼り付けてください。";

function selectCandidate(name, chip) {
  senderInput.value = name;
  for (const other of candidateBox.querySelectorAll(".chip")) {
    other.setAttribute("aria-pressed", String(other === chip));
  }
}

function showCandidates(candidates) {
  candidateBox.textContent = "";
  for (const candidate of candidates) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = candidate.name;
    chip.title = candidate.address ? `${candidate.name} <${candidate.address}>` : candidate.name;
    chip.setAttribute("aria-pressed", "false");
    chip.addEventListener("click", () => selectCandidate(candidate.name, chip));
    candidateBox.append(chip);
  }
  return candidateBox.querySelector(".chip");
}

// Outlook's markup changes often, so the detected name is only offered as a
// suggestion. The field stays editable and the paste route keeps working.
async function detectSender() {
  try {
    const result = await chrome.runtime.sendMessage({ type: "detect-sender" });
    if (!result?.ok) throw new Error(result?.error || "差出人を取得できませんでした。");
    const candidates = result.candidates || [];
    if (!candidates.length) {
      const reasons = [];
      if (!result.diagnostics?.readingPane) reasons.push("閲覧ウィンドウを特定できません");
      if (!result.diagnostics?.recipientRow) reasons.push("宛先行を特定できません");
      hint.textContent = reasons.length
        ? `差出人を検出できません(${reasons.join("・")})。${MANUAL_HINT}`
        : MANUAL_HINT;
      return;
    }
    const firstChip = showCandidates(candidates);
    if (result.confident && !senderInput.value) {
      selectCandidate(candidates[0].name, firstChip);
      hint.textContent = "差出人を自動入力しました。違っていれば候補を選ぶか、直接修正してください。";
      return;
    }
    hint.textContent = "差出人の候補です。正しいものを選ぶか、直接入力してください。";
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
    senderInput.dispatchEvent(new Event("input"));
    status.textContent = "差出人名を貼り付けました。";
    senderInput.focus();
  } catch (error) {
    status.className = "error";
    status.textContent = `貼り付けできませんでした: ${error.message || "クリップボードを確認してください。"}`;
  }
});

senderInput.addEventListener("input", () => {
  for (const chip of candidateBox.querySelectorAll(".chip")) {
    chip.setAttribute("aria-pressed", String(chip.textContent === senderInput.value));
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
