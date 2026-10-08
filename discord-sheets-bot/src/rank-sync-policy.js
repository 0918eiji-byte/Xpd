/**
 * Decide the source of truth for a rank mismatch.
 *
 * A sheet change is intentionally applied to Discord only through the
 * explicit action checkbox.  This keeps a stale or partially edited sheet
 * row from changing an unrelated member during the polling pass.
 */
export function rankSyncDecision({ triggered, sheetRank, discordRank }) {
  const sheet = String(sheetRank || "").trim();
  const discord = String(discordRank || "").trim();
  if (triggered) return "sheet-to-discord";
  if (!sheet || sheet === discord) return "none";
  return "discord-to-sheet";
}
