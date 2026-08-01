import assert from "node:assert/strict";
import { buildFilename, isSupportedOutlookUrl, sanitizeFilenamePart, selectOutlookTab, stripReplyPrefix } from "../lib.js";

assert.equal(stripReplyPrefix("RE: Re[2]: FW： 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("返信: 転送：確認"), "確認");
assert.equal(sanitizeFilenamePart('A/B: C*?'), "A_B_ C__");
assert.equal(buildFilename({ sender: "山田 太郎", date: new Date("2026-07-16T16:13:00+09:00"), subject: "RE: 進捗/報告" }), "📩_山田 太郎_20260716_0713_進捗_報告.pdf");
assert.equal(isSupportedOutlookUrl("https://outlook.cloud.microsoft/mail/deeplink/read/AAkALgAAAA?ItemID=AAkALg%2FA&exvsurl=1"), true);
assert.equal(isSupportedOutlookUrl("https://outlook.cloud.microsoft.evil.example/mail"), false);
assert.equal(selectOutlookTab([
  { id: 1, active: true, url: "chrome-extension://another-extension/popup.html", lastAccessed: 30 },
  { id: 2, active: false, url: "https://outlook.cloud.microsoft/mail/deeplink/read/id", lastAccessed: 20 }
]).id, 2);
console.log("lib tests passed");
