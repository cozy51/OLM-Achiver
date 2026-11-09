import { listAttachments } from "./attachments.js";
import { buildFilename, isSupportedOutlookUrl, selectOutlookTab } from "./lib.js";
import { buildMarkdownDocument, extractBodyMarkdown, markdownDataUrl, markdownFilename } from "./markdown.js";
import { buildPrintPageExpression, outlookPrintMode } from "./print-page.js";
import { collectAttachmentTiles } from "./scan-attachments.js";
import { collectSenderCandidates } from "./scan-sender.js";
import { CONFIDENT_SENDER_SCORE, rankSenders } from "./sender.js";

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

// Read separately from the print page so a message can still be saved when the
// attachment well is built in a way the scan does not recognise: the file list
// is an addition to the PDF, never a reason to fail the save. The well count
// travels with the files so an empty list is never reported as silence when
// there was something there to read.
async function readAttachments(tabId) {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: collectAttachmentTiles,
      world: "ISOLATED"
    });
    const scan = results?.[0]?.result;
    return { files: listAttachments(scan), wells: scan?.wells || 0 };
  } catch {
    return { files: [], wells: 0 };
  }
}

// Read before the print copy is placed, from the live message. A body that
// cannot be read still leaves a Markdown file with the header and the
// attachment list, and the PDF is never held back by it.
async function readBodyMarkdown(tabId) {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: extractBodyMarkdown,
      world: "ISOLATED"
    });
    return results?.[0]?.result?.markdown || "";
  } catch {
    return "";
  }
}

// The Markdown goes straight to the downloads folder: the PDF's save dialog is
// where the user picks a place, and a second dialog for every mail would make
// the option a chore. A failure is reported beside the PDF result, not as one.
async function saveMarkdown({ pdfFilename, sender, metadata, body }) {
  const filename = markdownFilename(pdfFilename);
  try {
    const text = buildMarkdownDocument({
      subject: metadata.subject,
      sender,
      date: metadata.date,
      url: metadata.url,
      attachments: metadata.attachments,
      body
    });
    await chrome.downloads.download({
      url: markdownDataUrl(text),
      filename,
      saveAs: false,
      conflictAction: "uniquify"
    });
    return { markdownFilename: filename };
  } catch (error) {
    return { markdownError: error?.message || String(error) };
  }
}

async function saveWithPrintDialog(tab, senderName, wantsMarkdown) {
  const attachments = await readAttachments(tab.id);
  const body = wantsMarkdown ? await readBodyMarkdown(tab.id) : "";
  const startResults = await runStep("readMail", () => chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: outlookPrintMode,
    args: ["start", tab.url, attachments.files],
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
    date: prepared.metadata.date,
    attachments: prepared.metadata.attachments
  });
  // Saved before the print preview opens, which holds the page until closed.
  const markdown = wantsMarkdown
    ? await saveMarkdown({ pdfFilename: filename, sender: senderName, metadata: prepared.metadata, body })
    : {};
  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: outlookPrintMode,
    args: ["print", filename],
    world: "ISOLATED"
  });
  return {
    filename,
    printDialog: true,
    attachments: prepared.metadata.attachments || [],
    attachmentWells: attachments.wells,
    ...markdown
  };
}

// Reading the sender only needs the DOM, so it uses chrome.scripting instead of
// the debugger API: the popup can offer a candidate without Edge showing the
// "デバッガーが接続しました" bar before the user has asked to save anything.
async function detectSender() {
  const tab = await runStep("findTab", findOutlookTab);
  if (!tab) throw new Error("このウィンドウで Outlook on the web のメールを開いてください。");
  const results = await runStep("readMail", () => chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: collectSenderCandidates,
    world: "ISOLATED"
  }));
  const scan = results?.[0]?.result;
  if (!scan?.ok) {
    throw new Error(`【失敗箇所: ${STEP_LABELS.readMail}】${scan?.error || "Outlook画面から差出人を読み取れませんでした。"}`);
  }
  // Only the winner is sent on. Runners-up are still needed while ranking —
  // agreement between them is what decides confidence — but showing them would
  // just ask the user to re-judge a decision the ranking already made.
  const [best] = rankSenders(scan);
  return {
    sender: best || null,
    confident: Boolean(best && best.score >= CONFIDENT_SENDER_SCORE),
    // Lets the popup say which part of the page could not be identified
    // instead of only reporting that detection produced nothing.
    diagnostics: {
      readingPane: scan.readingPane,
      recipientRow: scan.recipientRowFound,
      scanned: scan.candidates.length
    }
  };
}

const HANDLERS = {
  "save-current-mail": request => saveCurrentMail(request.senderName, request.markdown === true),
  "detect-sender": () => detectSender()
};

chrome.runtime.onMessage.addListener((request, _sender, respond) => {
  const handler = HANDLERS[request?.type];
  if (!handler) return;
  handler(request).then(result => respond({ ok: true, ...result })).catch(error => respond({ ok: false, error: error.message }));
  return true;
});

async function saveCurrentMail(senderName, wantsMarkdown = false) {
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
      return saveWithPrintDialog(tab, senderName.trim(), wantsMarkdown);
    }
    const attachments = await readAttachments(tab.id);
    const body = wantsMarkdown ? await readBodyMarkdown(tab.id) : "";
    const prepared = await runStep("readMail", async () => {
      const evaluated = await chrome.debugger.sendCommand(target, "Runtime.evaluate", {
        expression: buildPrintPageExpression("start", tab.url, attachments.files),
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
      date: prepared.metadata.date,
      attachments: prepared.metadata.attachments
    });
    await runStep("download", () => chrome.downloads.download({
      url: `data:application/pdf;base64,${pdf.data}`,
      filename,
      saveAs: true
    }));
    const markdown = wantsMarkdown
      ? await saveMarkdown({ pdfFilename: filename, sender: senderName.trim(), metadata: prepared.metadata, body })
      : {};
    return { filename, attachments: prepared.metadata.attachments || [], attachmentWells: attachments.wells, ...markdown };
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
