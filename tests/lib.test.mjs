import assert from "node:assert/strict";
import { buildFilename, isSupportedOutlookUrl, parseMailDate, sanitizeFilenamePart, selectOutlookTab, stripReplyPrefix } from "../lib.js";
import { buildPrintPageExpression } from "../print-page.js";

assert.equal(stripReplyPrefix("RE: Re[2]: FW： 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("返信: 転送：確認"), "確認");
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
assert.match(expression, /searchRoot\.documentElement \|\| searchRoot/);
assert.match(expression, /このメッセージ/);
assert.match(expression, /score >= 4/);
assert.match(expression, /keyboard shortcuts/);
assert.match(expression, /recipientsTop/);
assert.match(expression, /\\uE000-\\uF8FF/);
console.log("lib tests passed");
