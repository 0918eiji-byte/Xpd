import assert from "node:assert/strict";
import { rankSyncDecision } from "../src/rank-sync-policy.js";

// Regression: an unchecked stale sheet rank must not mutate Discord roles.
assert.equal(rankSyncDecision({ triggered: false, sheetRank: "巡査", discordRank: "警部" }), "discord-to-sheet");
// Explicit execution is the only sheet -> Discord path.
assert.equal(rankSyncDecision({ triggered: true, sheetRank: "警部", discordRank: "巡査" }), "sheet-to-discord");
assert.equal(rankSyncDecision({ triggered: false, sheetRank: "警部", discordRank: "警部" }), "none");
assert.equal(rankSyncDecision({ triggered: false, sheetRank: "", discordRank: "警部" }), "none");
console.log("rank-sync-policy: ok");
