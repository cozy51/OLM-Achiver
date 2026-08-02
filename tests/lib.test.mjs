import assert from "node:assert/strict";
import { buildFilename, isSupportedOutlookUrl, parseMailDate, sanitizeFilenamePart, selectOutlookTab, stripReplyPrefix } from "../lib.js";
import { buildPrintPageExpression } from "../print-page.js";
import { collectSenderCandidates } from "../scan-sender.js";
import { CONFIDENT_SENDER_SCORE, cleanSenderName, isSenderLike, rankSenders } from "../sender.js";
import { installActionIcon } from "../icon.js";

assert.equal(stripReplyPrefix("RE: Re[2]: FW： 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("返信: 転送：確認"), "確認");
assert.equal(stripReplyPrefix("[!]RE: 【依頼】確認"), "【依頼】確認");
assert.equal(stripReplyPrefix("[!]RE_ FW- 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("【依頼】RE: 確認"), "【依頼】RE: 確認");
assert.equal(sanitizeFilenamePart('A/B: C*?'), "A_B_ C__");
assert.equal(buildFilename({ sender: "山田 太郎", date: new Date("2026-07-16T16:13:00+09:00"), subject: "RE: 進捗/報告" }), "📩_山田 太郎_20260716_0713_進捗_報告.pdf");
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
assert.equal(typeof installActionIcon, "function");

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

console.log("lib tests passed");
