import { buildFilename, isSupportedOutlookUrl, selectOutlookTab } from "./lib.js";
import { buildPrintPageExpression } from "./print-page.js";

const STEP_LABELS = {
  findTab: "Outlook タブの特定",
  attach: "Edge のPDF機能への接続",
  readMail: "メール内容とURLの取得",
  createPdf: "PDFの作成",
  download: "PDFの保存"
};

async function runStep(step, operation) {
  try {
    return await operation();
  } catch (error) {
    const detail = error?.message || String(error);
    throw new Error(`【失敗箇所: ${STEP_LABELS[step]}】${detail}`);
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
  const tab = await runStep("findTab", findOutlookTab);
  if (!tab) {
    throw new Error("このウィンドウで Outlook on the web のメールを開いてください。");
  }

  const target = { tabId: tab.id };
  try {
    await runStep("attach", async () => {
      // Re-read the URL just before attaching, because a tab can navigate
      // between selection and debugger attachment.
      const current = await chrome.tabs.get(tab.id);
      if (!isSupportedOutlookUrl(current.url)) throw new Error("対象タブがOutlook以外のページへ移動しました。");
      await chrome.debugger.attach(target, "1.3");
      await chrome.debugger.sendCommand(target, "Page.enable");
    });
    const prepared = await runStep("readMail", async () => {
      const evaluated = await chrome.debugger.sendCommand(target, "Runtime.evaluate", {
        expression: buildPrintPageExpression("start", tab.url),
        returnByValue: true,
        awaitPromise: true
      });
      if (evaluated.exceptionDetails) throw new Error(evaluated.exceptionDetails.text || "メールの読み取りに失敗しました。");
      const result = evaluated.result?.value;
      if (!result?.ok) throw new Error(result?.error || "メールの読み取りに失敗しました。");
      return result;
    });
    const pdf = await runStep("createPdf", () => chrome.debugger.sendCommand(target, "Page.printToPDF", {
        printBackground: true,
        preferCSSPageSize: true,
        displayHeaderFooter: false
      })
    );
    const filename = buildFilename({
      sender: prepared.metadata.sender,
      subject: prepared.metadata.subject,
      date: new Date(prepared.metadata.date)
    });
    await runStep("download", () => chrome.downloads.download({
      url: `data:application/pdf;base64,${pdf.data}`,
      filename,
      saveAs: true
    }));
    return { filename };
  } finally {
    await chrome.debugger.sendCommand(target, "Runtime.evaluate", {
      expression: buildPrintPageExpression("end"),
      returnByValue: true
    }).catch(() => {});
    await chrome.debugger.detach(target).catch(() => {});
  }
}
