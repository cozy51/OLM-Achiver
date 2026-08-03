import assert from "node:assert/strict";
import { buildFilename, describeAttachments, isSupportedOutlookUrl, parseMailDate, sanitizeFilenamePart, selectOutlookTab, stripReplyPrefix } from "../lib.js";
import { buildPrintPageExpression } from "../print-page.js";
import { attachmentLink, attachmentNameFrom, listAttachments } from "../attachments.js";
import { collectAttachmentTiles } from "../scan-attachments.js";
import { collectSenderCandidates } from "../scan-sender.js";
import { CONFIDENT_SENDER_SCORE, cleanSenderName, isSenderLike, rankSenders } from "../sender.js";
import { installActionIcon } from "../icon.js";

assert.equal(stripReplyPrefix("RE: Re[2]: FW： 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("返信: 転送：確認"), "確認");
assert.equal(stripReplyPrefix("[!]RE: 【依頼】確認"), "【依頼】確認");
assert.equal(stripReplyPrefix("[!]RE_ FW- 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("【依頼】RE: 確認"), "【依頼】RE: 確認");
assert.equal(sanitizeFilenamePart('A/B: C*?'), "A-B- C--");
// A sender such as "Ito Takeshi / 伊藤剛" must stay one readable field.
assert.equal(sanitizeFilenamePart("Ito Takeshi / 伊藤剛"), "Ito Takeshi - 伊藤剛");
assert.equal(buildFilename({ sender: "山田 太郎", date: new Date("2026-07-16T16:13:00+09:00"), subject: "RE: 進捗/報告" }), "📩_山田 太郎_20260716_0713_進捗-報告.pdf");
assert.equal(buildFilename({ sender: "渡邉 陽介", date: "2026-07-13 (月) 20:41", subject: "【依頼】確認" }), "📩_渡邉 陽介_20260713_2041_【依頼】確認.pdf");
assert.equal(parseMailDate("2026年7月13日 20:41").getMinutes(), 41);
assert.equal(buildFilename({ sender: "山田", date: "", subject: "確認" }), "📩_山田_受信日不明_時刻不明_確認.pdf");
assert.equal(isSupportedOutlookUrl("https://outlook.cloud.microsoft/mail/deeplink/read/AAkALgAAAA?ItemID=AAkALg%2FA&exvsurl=1"), true);
assert.equal(isSupportedOutlookUrl("https://outlook.cloud.microsoft.evil.example/mail"), false);
assert.equal(selectOutlookTab([
  { id: 1, active: true, url: "chrome-extension://another-extension/popup.html", lastAccessed: 30 },
  { id: 2, active: false, url: "https://outlook.cloud.microsoft/mail/deeplink/read/id", lastAccessed: 20 }
]).id, 2);
const expression = buildPrintPageExpression("start", 'https://outlook.cloud.microsoft/mail/deeplink/read/a?x="quoted"');
assert.match(expression, /outlookPrintMode/);
assert.match(expression, /\\"quoted\\"/);
assert.doesNotMatch(expression, /findSender/);
assert.match(expression, /findMessageUrl/);
assert.match(expression, /aria-selected/);
assert.match(expression, /mail\/inbox\/id/);
assert.match(expression, /toCloudDeepLink/);
assert.match(expression, /outlook\.cloud\.microsoft/);
assert.match(expression, /exvsurl=1/);
assert.match(expression, /olm-attachments/);
assert.equal(typeof installActionIcon, "function");

// The attachment list travels as an argument, so the page renders exactly what
// was collected instead of scanning the DOM a second time.
const withAttachments = buildPrintPageExpression("start", "https://outlook.cloud.microsoft/mail/deeplink/read/a", [
  { name: "見積書.pdf", url: "https://contoso.sharepoint.com/見積書.pdf" }
]);
assert.match(withAttachments, /見積書\.pdf/);
assert.match(withAttachments, /sharepoint/);

assert.equal(describeAttachments([]), "");
assert.equal(describeAttachments(undefined), "");
// A message with nothing attached says nothing; a well that was found but
// could not be read must not look the same to the user.
assert.equal(describeAttachments([], { wells: 0 }), "");
assert.match(describeAttachments([], { wells: 1 }), /読み取れませんでした/);
assert.match(describeAttachments([{ name: "a.pdf" }], { wells: 1 }), /^添付ファイル 1件/);
assert.equal(
  describeAttachments([{ name: "見積書.pdf" }, { name: "写真.jpg" }]),
  "添付ファイル 2件の名前とリンクをPDFに記載しました: 見積書.pdf / 写真.jpg"
);
// Long lists stay one readable line; the PDF holds the complete list.
assert.match(describeAttachments([{ name: "a.pdf" }, { name: "b.pdf" }, { name: "c.pdf" }, { name: "d.pdf" }]), / ほか1件$/);

assert.equal(cleanSenderName("差出人: 山田 太郎 (Yamada,T.)"), "山田 太郎 (Yamada,T.)");
assert.equal(cleanSenderName("山田 太郎 <yamada@example.com> のプロフィール カードを開く"), "山田 太郎 の");
assert.equal(cleanSenderName("yamada@example.com"), "");
assert.equal(isSenderLike("宛先: 鈴木 花子"), false);
assert.equal(isSenderLike("全員に返信"), false);
assert.equal(isSenderLike("2026年7月13日 20:41"), false);
assert.equal(isSenderLike("Ctrl + R"), false);
assert.equal(isSenderLike("山田 太郎"), true);

// The sender header, the toolbar, the recipient row and the selected mail row
// as Outlook renders them together. The toolbar and the recipient used to win.
const scan = {
  recipientsTop: 160,
  candidates: [
    { source: "header", raw: "全員に返信", top: 100, bottom: 120 },
    { source: "header", raw: "Ctrl + Shift + R", top: 100, bottom: 120 },
    { source: "header", raw: "山田 太郎 (Yamada,T.) <yamada@example.com>", top: 130, bottom: 150 },
    { source: "header", raw: "宛先: 鈴木 花子", top: 160, bottom: 180 },
    { source: "header", raw: "鈴木 花子 (Suzuki,H.)", top: 161, bottom: 181 },
    { source: "selected-row", raw: "山田 太郎 (Yamada,T.)", order: 0 },
    { source: "selected-row", raw: "RE: 月次報告", order: 1 }
  ]
};
const ranked = rankSenders(scan);
assert.equal(ranked[0].name, "山田 太郎 (Yamada,T.)");
assert.equal(ranked[0].address, "yamada@example.com");
// Header and list row agreeing is what makes automatic filling trustworthy.
assert.deepEqual(ranked[0].sources, ["header", "selected-row"]);
assert.ok(ranked[0].score >= CONFIDENT_SENDER_SCORE);
assert.ok(!ranked.some(entry => /返信|宛先|Ctrl/.test(entry.name)));
assert.ok(!ranked.some(entry => entry.name === "RE: 月次報告"));
// A recipient below the "宛先:" row must never outrank the sender.
assert.ok(ranked.findIndex(entry => entry.name === "鈴木 花子 (Suzuki,H.)") > 0);

// Toolbar commands read exactly like plain names, so they are only kept out by
// the missing person evidence. These three were offered as senders in practice.
const toolbar = rankSenders({
  recipientsTop: 200,
  candidates: [
    { source: "header", raw: "More items", top: 10, bottom: 30 },
    { source: "header", raw: "アプリ", top: 10, bottom: 30 },
    { source: "header", raw: "リアクション", top: 10, bottom: 30 }
  ]
});
assert.deepEqual(toolbar, []);
// The same header text is accepted once Outlook marks it as a profile card.
assert.equal(rankSenders({
  candidates: [{ source: "header", raw: "原田 清枝", persona: true, top: 10, bottom: 30 }]
})[0].name, "原田 清枝");

// Taken from a real outlook.cloud.microsoft deeplink read view: the sender
// heading is "MSG_<item>_FROM", the display name uses full-width brackets, and
// the recipient well starts 3px below the sender row.
const realView = rankSenders({
  recipientsTop: 240,
  candidates: [
    { source: "explicit", raw: "原田 清枝（Harada,K.）", persona: true, top: 217, bottom: 237 },
    { source: "explicit", raw: "差出人: 原田 清枝（Harada,K.）", persona: true, top: 217, bottom: 237 },
    { source: "header", raw: "原田 清枝（Harada,K.）, 送信済み", persona: true, top: 221, bottom: 240 },
    { source: "header", raw: "アプリ", top: 60, bottom: 80 },
    { source: "header", raw: "リアクション", top: 60, bottom: 80 }
  ]
});
assert.equal(realView[0].name, "原田 清枝（Harada,K.）");
assert.ok(realView[0].score >= CONFIDENT_SENDER_SCORE);
assert.equal(realView.length, 1);

// A lone weak guess stays a suggestion instead of being filled in silently.
const weak = rankSenders({
  recipientsTop: 40,
  candidates: [{ source: "header", raw: "山田", persona: true, top: 10, bottom: 30 }]
});
assert.equal(weak[0].name, "山田");
assert.ok(weak[0].score < CONFIDENT_SENDER_SCORE);
assert.deepEqual(rankSenders({ candidates: [] }), []);

// The scan runs inside the page, so it must survive serialisation unchanged.
const scanSource = collectSenderCandidates.toString();
assert.doesNotMatch(scanSource, /\bimport\b/);
assert.match(scanSource, /aria-selected/);
assert.match(scanSource, /宛先/);

const attachmentScanSource = collectAttachmentTiles.toString();
assert.doesNotMatch(attachmentScanSource, /\bimport\b/);
// Outlook ids the well "MSG_<item>_ATTACHMENTS"; the label is the fallback.
assert.match(attachmentScanSource, /_ATTACHMENTS/);
assert.match(attachmentScanSource, /添付ファイル/);
// An image attachment shows a thumbnail and no name, so the name has to be
// looked for everywhere the accessibility tree can carry it.
assert.match(attachmentScanSource, /aria-labelledby/);
assert.match(attachmentScanSource, /"alt"/);
assert.match(attachmentScanSource, /"download"/);
// The well is looked for in the whole document when the pane holds none.
assert.match(attachmentScanSource, /findWells\(document\)/);
// A tile resolves to its [role="option"] before any inner draggable preview,
// so one attachment is reported once.
assert.match(attachmentScanSource, /closest\('\[role="option"\], \[role="listitem"\]'\)/);

// A name reached only through an image's alt text or a hidden label still ends
// up in the list; the href of a cloud tile must never become the name.
assert.deepEqual(listAttachments({
  pageUrl: "https://outlook.cloud.microsoft/mail/x",
  entries: [
    { tile: 0, labels: ["画像添付ファイル"], href: "" },
    { tile: 1, labels: ["写真.jpg"], href: "https://contoso.sharepoint.com/写真.jpg" }
  ]
}), [{ name: "写真.jpg", url: "https://contoso.sharepoint.com/写真.jpg" }]);

// Taken from a real outlook.cloud.microsoft read view: the tile's accessible
// name puts the file name first and joins the rest with spaces, so nothing but
// the extension marks where the name ends.
assert.equal(attachmentNameFrom("Image (1).jpg 開く 6 MB"), "Image (1).jpg");
assert.equal(attachmentNameFrom("見積書.xlsx 開く 2 MB"), "見積書.xlsx");
assert.equal(attachmentNameFrom("画像添付ファイル、写真.jpg、6 MB"), "写真.jpg");
assert.equal(attachmentNameFrom("見積書.xlsx (2 MB)"), "見積書.xlsx");
assert.equal(attachmentNameFrom("添付ファイル: 議事録.docx"), "議事録.docx");
// A comma inside the name itself must not split it into pieces.
assert.equal(attachmentNameFrom("報告書, 最終版.pdf"), "報告書, 最終版.pdf");
// Neither the well's own summary line nor its commands name a file.
assert.equal(attachmentNameFrom("添付ファイル 2 個 (6 MB)"), "");
assert.equal(attachmentNameFrom("すべてダウンロード"), "");
assert.equal(attachmentNameFrom("すべてを OneDrive - 村田機械株式会社 に保存"), "");
assert.equal(attachmentNameFrom(""), "");
// A size and a version number both look like a name plus an extension.
assert.equal(attachmentNameFrom("6.5MB"), "");
assert.equal(attachmentNameFrom("1.5 MB"), "");
// A URL in a data attribute is not a file name, however it ends.
assert.equal(attachmentNameFrom("https://contoso.sharepoint.com/thumb.png"), "");

// Only a real file address counts; an anchor back to the open message does not.
assert.equal(attachmentLink("https://contoso.sharepoint.com/a.pdf", "https://outlook.cloud.microsoft/mail/x"), "https://contoso.sharepoint.com/a.pdf");
assert.equal(attachmentLink("https://outlook.cloud.microsoft/mail/x#top", "https://outlook.cloud.microsoft/mail/x"), "");
assert.equal(attachmentLink("javascript:void(0)", "https://outlook.cloud.microsoft/mail/x"), "");

// The two image tiles of a real message: the name is on the [role="option"] and
// the preview inside it repeats it, which is one file, not two.
assert.deepEqual(listAttachments({
  pageUrl: "https://outlook.cloud.microsoft/mail/x",
  entries: [
    { tile: 0, labels: ["画像添付ファイル"], href: "" },
    { tile: 1, labels: ["Image (1).jpg 開く 6 MB"], href: "" },
    { tile: 1, labels: ["Image (1).jpg"], href: "" },
    { tile: 2, labels: ["Image (2).jpg 開く 6 MB"], href: "" },
    { tile: 3, labels: ["すべてダウンロード"], href: "" }
  ]
}), [
  { name: "Image (1).jpg", url: "" },
  { name: "Image (2).jpg", url: "" }
]);
// Two files really can share a name, so tiles are what separate them.
assert.equal(listAttachments({
  entries: [
    { tile: 0, labels: ["IMG_0001.jpg"], href: "" },
    { tile: 1, labels: ["IMG_0001.jpg"], href: "" }
  ]
}).length, 2);

const attachmentScan = {
  pageUrl: "https://outlook.cloud.microsoft/mail/x",
  entries: [
    { tile: 0, labels: ["添付ファイル 2 個 (6 MB)"], href: "" },
    // The name sits on the tile and the link on an element inside it.
    { tile: 1, labels: ["画像添付ファイル、写真.jpg、6 MB"], href: "" },
    { tile: 1, labels: ["写真.jpg"], href: "https://contoso.sharepoint.com/写真.jpg" },
    { tile: 2, labels: ["", "見積書.xlsx"], href: "" },
    { tile: 3, labels: ["すべてダウンロード"], href: "" }
  ]
};
assert.deepEqual(listAttachments(attachmentScan), [
  { name: "写真.jpg", url: "https://contoso.sharepoint.com/写真.jpg" },
  { name: "見積書.xlsx", url: "" }
]);
assert.deepEqual(listAttachments(undefined), []);
assert.deepEqual(listAttachments({ entries: [] }), []);

console.log("lib tests passed");
