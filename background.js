import { buildFilename } from "./lib.js";

const send = (tabId, payload) => chrome.tabs.sendMessage(tabId, payload);

chrome.runtime.onMessage.addListener((request, _sender, respond) => {
  if (request.type !== "save-current-mail") return;
  saveCurrentMail().then(result => respond({ ok: true, ...result })).catch(error => respond({ ok: false, error: error.message }));
  return true;
});

async function saveCurrentMail() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !/^https:\/\/outlook\.(office(?:365)?\.com|cloud\.microsoft)\//i.test(tab.url || "")) {
    throw new Error("Outlook on the web のメール画面を開いてください。");
  }

  let prepared = await send(tab.id, { type: "olm-pdf-print-mode", action: "start", url: tab.url });
  if (!prepared?.ok) throw new Error(prepared?.error || "メールの読み取りに失敗しました。");
  const target = { tabId: tab.id };
  try {
    await chrome.debugger.attach(target, "1.3");
    await chrome.debugger.sendCommand(target, "Page.enable");
    const pdf = await chrome.debugger.sendCommand(target, "Page.printToPDF", {
      printBackground: true,
      preferCSSPageSize: true,
      displayHeaderFooter: false
    });
    const filename = buildFilename({
      sender: prepared.metadata.sender,
      subject: prepared.metadata.subject,
      date: new Date(prepared.metadata.date)
    });
    await chrome.downloads.download({ url: `data:application/pdf;base64,${pdf.data}`, filename, saveAs: true });
    return { filename };
  } finally {
    await chrome.debugger.detach(target).catch(() => {});
    await send(tab.id, { type: "olm-pdf-print-mode", action: "end" }).catch(() => {});
  }
}
