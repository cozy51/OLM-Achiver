import { buildFilename, isSupportedOutlookUrl, selectOutlookTab } from "./lib.js";

const send = (tabId, payload) => chrome.tabs.sendMessage(tabId, payload);

async function preparePrintMode(tab) {
  const payload = { type: "olm-pdf-print-mode", action: "start", url: tab.url };
  try {
    return await send(tab.id, payload);
  } catch (error) {
    // A content script declared in the manifest is not added to tabs that were
    // already open when the extension was installed or reloaded. Inject it on
    // demand, then retry so users do not have to refresh Outlook themselves.
    if (!/Receiving end does not exist|Could not establish connection/i.test(error.message || "")) throw error;
    // Re-check the real tab URL immediately before injection. In some Edge
    // configurations the extension popup itself can temporarily be reported
    // as the active page.
    const currentTab = await chrome.tabs.get(tab.id);
    if (!isSupportedOutlookUrl(currentTab.url)) {
      throw new Error("Outlook のメールタブを特定できませんでした。メールを開いたタブに戻って再実行してください。");
    }
    await chrome.scripting.executeScript({
      target: { tabId: currentTab.id },
      files: ["content.js"],
      world: "ISOLATED"
    });
    return send(tab.id, payload);
  }
}

async function findOutlookTab() {
  const tabs = await chrome.tabs.query({ currentWindow: true });
  return selectOutlookTab(tabs);
}

chrome.runtime.onMessage.addListener((request, _sender, respond) => {
  if (request.type !== "save-current-mail") return;
  saveCurrentMail().then(result => respond({ ok: true, ...result })).catch(error => respond({ ok: false, error: error.message }));
  return true;
});

async function saveCurrentMail() {
  const tab = await findOutlookTab();
  if (!tab) {
    throw new Error("このウィンドウで Outlook on the web のメールを開いてください。");
  }

  const prepared = await preparePrintMode(tab);
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
