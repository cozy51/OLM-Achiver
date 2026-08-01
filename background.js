import { buildFilename, isSupportedOutlookUrl, selectOutlookTab } from "./lib.js";
import { buildPrintPageExpression } from "./print-page.js";

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

  const target = { tabId: tab.id };
  try {
    await chrome.debugger.attach(target, "1.3");
    await chrome.debugger.sendCommand(target, "Page.enable");
    const evaluated = await chrome.debugger.sendCommand(target, "Runtime.evaluate", {
      expression: buildPrintPageExpression("start", tab.url),
      returnByValue: true,
      awaitPromise: true
    });
    if (evaluated.exceptionDetails) throw new Error(evaluated.exceptionDetails.text || "メールの読み取りに失敗しました。");
    const prepared = evaluated.result?.value;
    if (!prepared?.ok) throw new Error(prepared?.error || "メールの読み取りに失敗しました。");
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
    await chrome.debugger.sendCommand(target, "Runtime.evaluate", {
      expression: buildPrintPageExpression("end"),
      returnByValue: true
    }).catch(() => {});
    await chrome.debugger.detach(target).catch(() => {});
  }
}
