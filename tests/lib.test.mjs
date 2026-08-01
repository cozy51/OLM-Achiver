import assert from "node:assert/strict";
import { buildFilename, sanitizeFilenamePart, stripReplyPrefix } from "../lib.js";

assert.equal(stripReplyPrefix("RE: Re[2]: FW： 月次報告"), "月次報告");
assert.equal(stripReplyPrefix("返信: 転送：確認"), "確認");
assert.equal(sanitizeFilenamePart('A/B: C*?'), "A_B_ C__");
assert.equal(buildFilename({ sender: "山田 太郎", date: new Date("2026-07-16T16:13:00+09:00"), subject: "RE: 進捗/報告" }), "📩_山田 太郎_20260716_0713_進捗_報告.pdf");
console.log("lib tests passed");
