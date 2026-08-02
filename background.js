import { buildFilename, isSupportedOutlookUrl, selectOutlookTab } from "./lib.js";
import { buildPrintPageExpression, outlookPrintMode } from "./print-page.js";
import { installActionIcon } from "./icon.js";

installActionIcon();
chrome.runtime.onInstalled.addListener(installActionIcon);
chrome.runtime.onStartup.addListener(installActionIcon);

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

async function saveWithPrintDialog(tab, senderName) {
  const startResults = await runStep("readMail", () => chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: outlookPrintMode,
    args: ["start", tab.url],
    world: "ISOLATED"
  }));
  const prepared = startResults?.[0]?.result;
  if (!prepared?.ok) {
    const detail = prepared?.error || (prepared === undefined
      ? "Outlook画面の解析結果を取得できませんでした。拡張機能を再読み込みして、メールを開き直してください。"
      : "メールの読み取りに失敗しました。");
    throw new Error(`【失敗箇所: メール内容とURLの取得】${detail}`);
  }
  const filename = buildFilename({
    sender: senderName,
    subject: prepared.metadata.subject,
    date: prepared.metadata.date
  });
  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: outlookPrintMode,
    args: ["print", filename],
    world: "ISOLATED"
  });
  return { filename, printDialog: true };
}

chrome.runtime.onMessage.addListener((request, _sender, respond) => {
  if (request.type !== "save-current-mail") return;
  saveCurrentMail(request.senderName).then(result => respond({ ok: true, ...result })).catch(error => respond({ ok: false, error: error.message }));
  return true;
});

async function saveCurrentMail(senderName) {
  if (!String(senderName || "").trim()) throw new Error("差出人名を入力してください。");
  const tab = await runStep("findTab", findOutlookTab);
  if (!tab) {
    throw new Error("このウィンドウで Outlook on the web のメールを開いてください。");
  }

  const target = { tabId: tab.id };
  let attached = false;
  try {
    // Re-read the URL because a tab can navigate between selection and use.
    const current = await chrome.tabs.get(tab.id);
    if (!isSupportedOutlookUrl(current.url)) {
      throw new Error(`【失敗箇所: ${STEP_LABELS.findTab}】対象タブがOutlook以外のページへ移動しました。`);
    }
    try {
      await chrome.debugger.attach(target, "1.3");
      attached = true;
      await chrome.debugger.sendCommand(target, "Page.enable");
    } catch (error) {
      // Some managed Edge environments prohibit debugger attachment. The
      // standard print dialog can still create a link-preserving PDF.
      if (attached) {
        await chrome.debugger.detach(target).catch(() => {});
        attached = false;
      }
      return saveWithPrintDialog(tab, senderName.trim());
    }
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
      sender: senderName.trim(),
      subject: prepared.metadata.subject,
      date: prepared.metadata.date
    });
    await runStep("download", () => chrome.downloads.download({
      url: `data:application/pdf;base64,${pdf.data}`,
      filename,
      saveAs: true
    }));
    return { filename };
  } finally {
    if (attached) {
      await chrome.debugger.sendCommand(target, "Runtime.evaluate", {
        expression: buildPrintPageExpression("end"),
        returnByValue: true
      }).catch(() => {});
      await chrome.debugger.detach(target).catch(() => {});
    }
  }
}
