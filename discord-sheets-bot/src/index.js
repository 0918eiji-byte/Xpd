import http from "node:http";
import { createHash } from "node:crypto";
import {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  Client,
  ChannelType,
  EmbedBuilder,
  GatewayIntentBits,
  MessageFlags,
  ModalBuilder,
  PermissionFlagsBits,
  SlashCommandBuilder,
  StringSelectMenuBuilder,
  StringSelectMenuOptionBuilder,
  TextInputBuilder,
  TextInputStyle,
} from "discord.js";
import { google } from "googleapis";

const required = [
  "DISCORD_BOT_TOKEN",
  "DISCORD_GUILD_ID",
  "GOOGLE_SPREADSHEET_ID",
  "GOOGLE_SERVICE_ACCOUNT_JSON",
];

const missing = required.filter((name) => !process.env[name]);
if (missing.length) throw new Error(`Railway Variables が未設定です: ${missing.join(", ")}`);
if (!process.env.ROSTER_ROLE_ID && !process.env.ROSTER_ROLE_IDS) {
  throw new Error("ROSTER_ROLE_ID または ROSTER_ROLE_IDS を設定してください。");
}

let credentials;
try {
  credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  credentials.private_key = credentials.private_key?.replace(/\\n/g, "\n");
} catch {
  throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON にJSONファイルの中身をそのまま貼り付けてください。");
}

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const sheets = google.sheets({ version: "v4", auth });
const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
const guildId = process.env.DISCORD_GUILD_ID;
const rosterRoleIds = new Set(
  (process.env.ROSTER_ROLE_IDS || process.env.ROSTER_ROLE_ID).split(",").map((id) => id.trim()).filter(Boolean),
);
const excludeRoleIds = new Set(
  (process.env.EXCLUDE_ROLE_IDS || "").split(",").map((id) => id.trim()).filter(Boolean),
);
const unifiedSettingsSheetName = "設定";
const unifiedSettingsSchema = "v2:READY";
const unifiedSettingsRanges = {
  rank: { range: "A10:F109", firstRow: 10 },
  bonus: { range: "A114:C213", firstRow: 114 },
  recruitment: { range: "A218:N317", firstRow: 218 },
  interview: { range: "A322:P421", firstRow: 322 },
  questions: { range: "A426:G1425", firstRow: 426 },
};
const unifiedSettingsBlocks = new Map([
  ["ランク", [7, 109]],
  ["ボーナス回", [111, 213]],
  ["書類選考", [215, 317]],
  ["面接", [319, 421]],
  ["面接質問", [423, 1425]],
  ["連携", [1427, 1437]],
  ["ランク報告", [1437, 1446]],
  ["操作ボード", [1447, 1455]],
]);
let unifiedSettingsSheetId = null;
let displayedSettingsCategory = "";
let requestedSettingsCategory = "すべて";
let lastSettingsViewCheckAt = 0;
let unifiedSettingsHealthy = false;
let lastUnifiedAuditAt = 0;
let unifiedSettingsSnapshot = null;
const recruitmentStatusCache = new Map();
const interviewStatusCache = new Map();
// Status cells are informational only.  Do not spend Sheets write quota on a
// heartbeat during every audit; enable explicitly when repairing the UI.
let lastUnifiedStatusKey = "";
let lastUnifiedStatusWriteAt = 0;

const unifiedSettingsCategoryOptions = [
  "すべて", "ランク", "ボーナス回", "書類選考", "面接", "面接質問", "連携", "ランク報告", "操作ボード",
];

async function ensureUnifiedSettingsCategoryValidation() {
  if (unifiedSettingsSheetId === null) {
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties(sheetId,title)",
    });
    unifiedSettingsSheetId = metadata.data.sheets
      ?.find((sheet) => sheet.properties?.title === unifiedSettingsSheetName)?.properties?.sheetId ?? null;
  }
  if (unifiedSettingsSheetId === null) throw new Error("設定シートが見つかりません");
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        setDataValidation: {
          range: { sheetId: unifiedSettingsSheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 1, endColumnIndex: 2 },
          rule: {
            condition: { type: "ONE_OF_LIST", values: unifiedSettingsCategoryOptions.map((value) => ({ userEnteredValue: value })) },
            strict: false,
            showCustomUi: true,
          },
        },
      }],
    },
  });
}

async function unifiedSettingsReady(force = false) {
  if (force) await auditUnifiedSettings(true);
  return Boolean(unifiedSettingsHealthy && unifiedSettingsSnapshot);
}

function normalizedId(value) {
  return String(value ?? "").replace(/^'+/, "").replace(/[\s`]/g, "").trim();
}

function isEnabledSetting(value) {
  if (value === true || value === 1) return true;
  const normalized = String(value ?? "").trim().toLowerCase();
  return ["true", "はい", "yes", "有効", "1", "✓"].includes(normalized);
}

function formatDiscordIdError(error, discordId = "") {
  const raw = String(error?.response?.data?.message || error?.message || error || "");
  const code = Number(error?.code || error?.rawError?.code || error?.response?.data?.code || 0);
  if ([10004, 10007, 10013].includes(code) || /unknown (member|user|guild)|unknown member|unknown user/i.test(raw)) {
    return `Discord IDが違います${discordId ? `（${discordId}）` : ""}。Discordサーバーに存在するユーザーIDを確認してください。`;
  }
  return raw;
}

function duplicateValues(values) {
  const seen = new Set();
  const duplicates = new Set();
  for (const value of values.filter(Boolean)) {
    if (seen.has(value)) duplicates.add(value);
    seen.add(value);
  }
  return [...duplicates];
}

async function auditUnifiedSettings(force = false) {
  const now = Date.now();
  if (!force && now - lastUnifiedAuditAt < 120000) return unifiedSettingsHealthy;
  lastUnifiedAuditAt = now;
  let schema = "";
  try {
    const response = await sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: [
        `'${unifiedSettingsSheetName}'!K3`,
        `'${unifiedSettingsSheetName}'!B3`,
        `'${unifiedSettingsSheetName}'!${unifiedSettingsRanges.rank.range}`,
        `'${unifiedSettingsSheetName}'!${unifiedSettingsRanges.bonus.range}`,
        `'${unifiedSettingsSheetName}'!${unifiedSettingsRanges.recruitment.range}`,
        `'${unifiedSettingsSheetName}'!${unifiedSettingsRanges.interview.range}`,
        `'${unifiedSettingsSheetName}'!${unifiedSettingsRanges.questions.range}`,
      ],
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    const ranges = response.data.valueRanges || [];
    schema = String(ranges[0]?.values?.[0]?.[0] || "").trim();
    if (schema !== unifiedSettingsSchema) return false;
    requestedSettingsCategory = String(ranges[1]?.values?.[0]?.[0] || "すべて").trim();
    const [rankRows, bonusRows, recruitmentRows, interviewRows, questionRows] = ranges.slice(2).map((range) => range.values || []);
    const errors = [];
    const snowflake = /^\d{17,20}$/;
    const activeRanks = rankRows.filter((row) => isEnabledSetting(row[5]));
    for (const row of activeRanks) {
      if (!String(row[1] || "").trim() || !snowflake.test(normalizedId(row[2]))) errors.push("ランク名またはロールIDが不正");
      if (!Number.isFinite(Number(row[0])) || !Number.isFinite(Number(row[4]))) errors.push("ランク優先度またはボーナス係数が不正");
    }
    if (duplicateValues(activeRanks.map((row) => String(row[1] || "").trim())).length) errors.push("ランク名が重複");
    if (duplicateValues(activeRanks.map((row) => normalizedId(row[2]))).length) errors.push("ランクロールIDが重複");
    if (duplicateValues(activeRanks.map((row) => String(row[0] || "").trim())).length) errors.push("ランク優先度が重複");
    for (const requiredRank of ["体験", "解雇者"]) {
      if (activeRanks.filter((row) => String(row[1] || "").trim() === requiredRank).length !== 1) errors.push(`${requiredRank}ランクは有効な1件が必要`);
    }
    const rounds = [];
    for (const row of recruitmentRows.filter((item) => String(item[0] || "").trim() === "はい")) {
      const round = String(row[1] || "").trim();
      rounds.push(round);
      if (!round || !String(row[2] || "").trim()) errors.push("書類選考の募集回または回答URLが未設定");
      else {
        try { extractSpreadsheetId(String(row[2])); } catch { errors.push("書類選考の回答スプレッドシートURLが不正"); }
      }
      for (const [value, label] of [[row[3], "投票"], [row[7], "合格発表"], [row[13], "書類合格ロール"]]) {
        if (normalizedId(value) && !snowflake.test(normalizedId(value))) errors.push(`書類選考の${label}IDが不正`);
      }
      if (!(Number(row[5]) >= 1 && Number(row[5]) <= 168)) errors.push("書類選考の期限は1～168時間");
      if (!(Number(row[6]) >= 1 && Number(row[6]) <= 1000)) errors.push("書類選考の締切票数は1～1000");
    }
    if (duplicateValues(rounds).length) errors.push("書類選考の募集回が重複");
    const interviewRounds = [];
    for (const row of interviewRows.filter((item) => String(item[0] || "").trim() === "はい")) {
      const round = String(row[1] || "").trim();
      interviewRounds.push(round);
      if (!round || !String(row[10] || "").trim()) errors.push("面接の募集回または質問セットが未設定");
      for (const [value, label] of [[row[2], "実行チャンネル"], [row[3], "面接官ロール"], [row[4], "投票チャンネル"], [row[8], "合格発表チャンネル"], [row[15], "面接合格ロール"]]) {
        if (normalizedId(value) && !snowflake.test(normalizedId(value))) errors.push(`面接の${label}IDが不正`);
      }
      if (!(Number(row[6]) >= 1 && Number(row[6]) <= 168)) errors.push("面接の期限は1～168時間");
      if (!(Number(row[7]) >= 1 && Number(row[7]) <= 1000)) errors.push("面接の締切票数は1～1000");
    }
    if (duplicateValues(interviewRounds).length) errors.push("面接の募集回が重複");
    const activeQuestions = questionRows.filter((row) => String(row[0] || "").trim() === "はい");
    if (duplicateValues(activeQuestions.map((row) => `${String(row[1] || "").trim()}:${String(row[2] || "").trim()}`)).length) errors.push("質問IDが質問セット内で重複");
    const questionSets = new Map();
    for (const row of activeQuestions) {
      const setName = String(row[1] || "").trim();
      if (!setName || !String(row[2] || "").trim() || !String(row[4] || "").trim()) errors.push("面接質問の必須項目が未設定");
      const list = questionSets.get(setName) || [];
      list.push(row);
      questionSets.set(setName, list);
    }
    for (const [setName, rows] of questionSets) {
      if (rows.length > 24) errors.push(`${setName}は24問以下にしてください`);
      if (!rows.some((row) => String(row[6] || "").trim() !== "いいえ")) errors.push(`${setName}に必須質問が必要`);
      if (duplicateValues(rows.map((row) => String(row[3] || "").trim())).length) errors.push(`${setName}の表示順が重複`);
    }
    for (const row of interviewRows.filter((item) => String(item[0] || "").trim() === "はい")) {
      const setName = String(row[10] || "").trim();
      if (!questionSets.has(setName)) errors.push(`面接で参照する質問セット「${setName}」がありません`);
    }
    const configuredBonusRows = bonusRows.filter((row) => String(row[0] || "").trim());
    if (duplicateValues(configuredBonusRows.map((row) => String(row[0] || "").trim())).length) errors.push("ボーナス支給回が重複");
    if (configuredBonusRows.some((row) => !Number.isFinite(Number(row[1])) || Number(row[1]) < 0)) errors.push("ボーナスのプール開始額が不正");
    const uniqueErrors = [...new Set(errors)];
    unifiedSettingsHealthy = uniqueErrors.length === 0;
    if (unifiedSettingsHealthy) {
      unifiedSettingsSnapshot = {
        rank: rankRows.map((row) => [...row]),
        bonus: bonusRows.map((row) => [...row]),
        recruitment: recruitmentRows.map((row) => [...row]),
        interview: interviewRows.map((row) => [...row]),
        questions: questionRows.map((row) => [...row]),
        auditedAt: now,
      };
    }
    const statusKey = uniqueErrors.length ? `要確認: ${uniqueErrors.slice(0, 3).join(" / ")}` : "正常";
    const shouldWriteStatus = process.env.SETTINGS_HEARTBEAT_WRITE === "1"
      && (statusKey !== lastUnifiedStatusKey || now - lastUnifiedStatusWriteAt >= 900000);
    if (shouldWriteStatus) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            { range: `'${unifiedSettingsSheetName}'!E3`, values: [[statusKey]] },
            { range: `'${unifiedSettingsSheetName}'!H3`, values: [[sheetDateTime(new Date())]] },
          ],
        },
      });
      lastUnifiedStatusKey = statusKey;
      lastUnifiedStatusWriteAt = now;
    }
    return unifiedSettingsHealthy;
  } catch (error) {
    if (isSheetsQuotaError(error)) lastUnifiedAuditAt = 0;
    if (!schema && (error.code === 400 || /Unable to parse range|not found/i.test(error.message || ""))) return false;
    throw error;
  }
}

async function unifiedRange(key) {
  if (!await unifiedSettingsReady()) return null;
  const block = unifiedSettingsRanges[key];
  const rows = unifiedSettingsSnapshot?.[key];
  if (!rows) return null;
  return { rows: rows.map((row) => [...row]), firstRow: block.firstRow, source: "unified" };
}

async function applyUnifiedSettingsView() {
  if (!await unifiedSettingsReady()) return;
  const now = Date.now();
  if (now - lastSettingsViewCheckAt >= 10000) {
    lastSettingsViewCheckAt = now;
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${unifiedSettingsSheetName}'!B3`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    requestedSettingsCategory = String(response.data.values?.[0]?.[0] || "すべて").trim();
  }
  const requested = requestedSettingsCategory;
  const category = requested === "すべて" || unifiedSettingsBlocks.has(requested) ? requested : "すべて";
  if (category === displayedSettingsCategory) return;
  if (unifiedSettingsSheetId === null) {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties(sheetId,title)" });
    unifiedSettingsSheetId = metadata.data.sheets
      ?.find((sheet) => sheet.properties?.title === unifiedSettingsSheetName)?.properties?.sheetId;
  }
  if (unifiedSettingsSheetId === undefined || unifiedSettingsSheetId === null) return;
  const requests = [...unifiedSettingsBlocks.entries()].map(([name, [startIndex, endIndex]]) => ({
    updateDimensionProperties: {
      range: {
        sheetId: unifiedSettingsSheetId,
        dimension: "ROWS",
        startIndex,
        endIndex,
      },
      properties: { hiddenByUser: category !== "すべて" && category !== name },
      fields: "hiddenByUser",
    },
  }));
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  displayedSettingsCategory = category;
}

const client = new Client({
  intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMembers, GatewayIntentBits.GuildMessages],
});

// Keep background Sheets work isolated from user-facing Discord operations.
// A slow polling/synchronization job must never delay an interaction response.
let sheetsQueue = Promise.resolve();
let discordOperationQueue = Promise.resolve();
const lastSynchronizedRanks = new Map();
let sheetsQuotaBackoffUntil = 0;
let sheetsQuotaBackoffLevel = 0;
let maintenanceQuotaBackoffUntil = 0;
const announcementRetryButtonsInstalled = new Set();

function isSheetsQuotaError(error) {
  const status = Number(error?.code || error?.response?.status || error?.response?.data?.error?.code || 0);
  const message = String(error?.response?.data?.error?.message || error?.message || error || "");
  return status === 429 || /quota exceeded|rate limit|read requests per minute|RESOURCE_EXHAUSTED/i.test(message);
}

function registerSheetsQuotaError(error) {
  if (!isSheetsQuotaError(error)) return false;
  sheetsQuotaBackoffLevel = Math.min(sheetsQuotaBackoffLevel + 1, 4);
  const delayMs = Math.min(60000 * (2 ** (sheetsQuotaBackoffLevel - 1)), 5 * 60000) + Math.floor(Math.random() * 5000);
  sheetsQuotaBackoffUntil = Math.max(sheetsQuotaBackoffUntil, Date.now() + delayMs);
  console.error(`Google Sheets読込制限: ${Math.ceil(delayMs / 1000)}秒後に自動再試行します`);
  return true;
}

function sheetsBackoffRemainingSeconds() {
  return Math.max(0, Math.ceil((sheetsQuotaBackoffUntil - Date.now()) / 1000));
}

function enqueueSheets(label, work) {
  sheetsQueue = sheetsQueue.then(work).catch((error) => {
    const message = error.response?.data?.error?.message || error.message || error;
    console.error(`[${label}]`, message);
  });
  return sheetsQueue;
}

function enqueueDiscordOperation(label, work) {
  discordOperationQueue = discordOperationQueue.then(work).catch((error) => {
    const message = error.response?.data?.error?.message || error.message || error;
    console.error(`[${label}]`, message);
  });
  return discordOperationQueue;
}

async function readRankMap() {
  const unified = await unifiedRange("rank");
  const response = unified || await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'ランク設定'!A3:F1000",
  });
  const map = new Map();
  for (const row of unified?.rows || response.data.values || []) {
    const priority = Number(row[0]);
    const rankName = String(row[1] || "").trim();
    const roleId = normalizedId(row[2]);
    const roleName = String(row[3] || rankName).trim();
    const enabled = isEnabledSetting(row[5]);
    if (enabled && roleId && rankName) map.set(roleId, { roleId, priority, rankName, roleName });
  }
  return map;
}

function assessMember(member, rankMap) {
  const matches = member.roles.cache
    .map((role) => rankMap.get(role.id))
    .filter((rank) => rank && rank.rankName !== "解雇者")
    .sort((a, b) => a.priority - b.priority);
  return {
    roleNames: [...new Set(matches.map((item) => item.roleName))].join(", "),
    rankName: matches[0]?.rankName || "",
    priority: matches[0]?.priority ?? 9999,
  };
}

function sortedRanks(rankMap) {
  return [...rankMap.values()]
    .filter((rank) => rank.rankName !== "解雇者")
    .sort((a, b) => a.priority - b.priority);
}

function dismissedRank(rankMap) {
  return [...rankMap.values()].find((rank) => rank.rankName === "解雇者") || null;
}

function rankByName(rankMap, rankName) {
  return [...rankMap.values()].find((rank) => rank.rankName === rankName) || null;
}

async function ensureRosterRole(guild, member, reason) {
  if ([...rosterRoleIds].some((roleId) => member.roles.cache.has(roleId))) return false;
  const roleId = [...rosterRoleIds][0];
  const role = guild.roles.cache.get(roleId);
  if (!role) throw new Error(`Police Officerロールが見つかりません: ${roleId}`);
  if (!role.editable) throw new Error(`Botより上位のロールは付与できません: ${role.name}`);
  await member.roles.add(roleId, reason);
  return true;
}

async function removeRosterRoles(guild, member, reason) {
  const roleIds = [...rosterRoleIds].filter((roleId) => member.roles.cache.has(roleId));
  if (!roleIds.length) return [];
  const roles = roleIds.map((roleId) => guild.roles.cache.get(roleId)).filter(Boolean);
  const missingIds = roleIds.filter((roleId) => !guild.roles.cache.has(roleId));
  if (missingIds.length) throw new Error(`Police Officerロールが見つかりません: ${missingIds.join(", ")}`);
  const blocked = roles.filter((role) => !role.editable).map((role) => role.name);
  if (blocked.length) throw new Error(`Botより上位のロールは解除できません: ${blocked.join(", ")}`);
  await member.roles.remove(roleIds, reason);
  return roles.map((role) => role.name);
}

async function removeAllEditableRoles(guild, member, reason) {
  const removableRoles = member.roles.cache.filter((role) =>
    role.id !== guild.id && !role.managed,
  );
  const blockedRoles = removableRoles.filter((role) => !role.editable);
  if (blockedRoles.size) {
    throw new Error(`Botより上位のため全ロールを解除できません: ${blockedRoles.map((role) => role.name).join(", ")}`);
  }
  if (removableRoles.size) await member.roles.remove([...removableRoles.keys()], reason);
  return removableRoles.map((role) => role.name);
}

function normalizedRoleIds(roleIds) {
  return [...new Set(roleIds.map((id) => String(id || "").replace(/^'/, "").trim()).filter(Boolean))];
}

function configuredRoleId(guild, configuredId, fallbackNames = []) {
  const explicitId = String(configuredId || "").replace(/^'/, "").trim();
  if (explicitId) return explicitId;
  const names = new Set(fallbackNames.map((name) => String(name).trim().toLowerCase()));
  const matches = guild.roles.cache.filter((role) => names.has(role.name.trim().toLowerCase()));
  if (matches.size > 1) {
    throw new Error(`同名候補のDiscordロールが複数あります: ${fallbackNames.join(" / ")}`);
  }
  return matches.first()?.id || "";
}

async function reconcileMemberRoles(guild, member, { add = [], remove = [] }, reason) {
  const addIds = normalizedRoleIds(add).filter((id) => !member.roles.cache.has(id));
  const removeIds = normalizedRoleIds(remove).filter((id) => member.roles.cache.has(id) && !addIds.includes(id));
  const allIds = [...new Set([...addIds, ...removeIds])];
  if (!allIds.length) return { added: [], removed: [] };
  if (!guild.members.me?.permissions.has(PermissionFlagsBits.ManageRoles)) {
    throw new Error("Botにロール管理権限がありません");
  }
  const missing = allIds.filter((id) => !guild.roles.cache.has(id));
  if (missing.length) throw new Error(`Discordに存在しないロールIDです: ${missing.join(", ")}`);
  const blocked = allIds
    .map((id) => guild.roles.cache.get(id))
    .filter((role) => role.managed || !role.editable)
    .map((role) => role.name);
  if (blocked.length) throw new Error(`Botより上位または管理対象のため操作できないロールです: ${blocked.join(", ")}`);

  // 次段階のロールが付かなかった場合に、前段階ロールだけ外れることを防ぐ。
  if (addIds.length) await member.roles.add(addIds, reason);
  if (removeIds.length) await member.roles.remove(removeIds, reason);
  return {
    added: addIds.map((id) => guild.roles.cache.get(id)?.name || id),
    removed: removeIds.map((id) => guild.roles.cache.get(id)?.name || id),
  };
}

async function applySelectedRank(guild, member, rankMap, requestedRank, reason) {
  const selected = String(requestedRank || "").trim();
  const ranks = sortedRanks(rankMap);
  const currentRanks = ranks.filter((rank) => member.roles.cache.has(rank.roleId));
  const currentName = currentRanks[0]?.rankName || "階級なし";
  const dismissal = dismissedRank(rankMap);

  if (selected === "解雇" || selected === "解雇者") {
    if (!dismissal) throw new Error("ランク設定に有効な「解雇者」ロールがありません");
    const dismissalRole = guild.roles.cache.get(dismissal.roleId);
    if (!dismissalRole) throw new Error(`解雇者ロールが見つかりません: ${dismissal.roleId}`);
    if (!dismissalRole.editable) throw new Error(`Botより上位のロールは付与できません: ${dismissalRole.name}`);
    const removeIds = [...new Set([
      ...currentRanks.map((rank) => rank.roleId),
      ...[...rosterRoleIds].filter((roleId) => member.roles.cache.has(roleId)),
    ])];
    const blocked = removeIds
      .map((roleId) => guild.roles.cache.get(roleId))
      .filter((role) => role && !role.editable)
      .map((role) => role.name);
    if (blocked.length) throw new Error(`Botより上位のロールは解除できません: ${blocked.join(", ")}`);
    if (!member.roles.cache.has(dismissal.roleId)) await member.roles.add(dismissal.roleId, reason);
    if (removeIds.length) await member.roles.remove(removeIds, reason);
    lastSynchronizedRanks.set(member.id, "解雇");
    return { transition: `${currentName} → 解雇者`, previousRank: currentRanks[0]?.rankName || "" };
  }

  const target = selected === "？？？？" ? null : rankByName(rankMap, selected);
  if (!target && selected !== "？？？？") throw new Error(`ランク設定にない階級です: ${selected || "未選択"}`);
  if ([...excludeRoleIds].some((roleId) => member.roles.cache.has(roleId))) {
    throw new Error("除外ロールが付いているためランクを変更できません");
  }
  await ensureRosterRole(guild, member, `${reason}: Police Officerを自動復元`);

  const targetRole = target ? guild.roles.cache.get(target.roleId) : null;
  if (target && !targetRole) throw new Error(`対象ロールが見つかりません: ${target.rankName}`);
  if (targetRole && !member.roles.cache.has(target.roleId) && !targetRole.editable) {
    throw new Error(`Botより上位のロールは操作できません: ${targetRole.name}`);
  }
  const removeIds = currentRanks.filter((rank) => rank.roleId !== target?.roleId).map((rank) => rank.roleId);
  const blocked = removeIds
    .map((roleId) => guild.roles.cache.get(roleId))
    .filter((role) => role && !role.editable)
    .map((role) => role.name);
  if (blocked.length) throw new Error(`Botより上位のロールは解除できません: ${blocked.join(", ")}`);
  if (removeIds.length) await member.roles.remove(removeIds, reason);
  if (targetRole && !member.roles.cache.has(target.roleId)) await member.roles.add(target.roleId, reason);
  const nextName = target?.rankName || "？？？？";
  lastSynchronizedRanks.set(member.id, nextName);
  return { transition: `${currentName} → ${nextName}`, previousRank: currentRanks[0]?.rankName || "" };
}

function columnLetter(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

const rankSelectionHeader = "変更後ランク";
const actionTriggerHeader = "実行ボタン";
const legacyActionTriggerHeader = "実行";
const sortPriorityHeader = "階級順序";
const bonusFactorHeader = "ランク係数";
const legacyBonusFactorHeader = "固定係数";
const currentStatusHeader = "現在状態";
const botHeaders = ["社員ID", "表示名", "Discordロール", "適用ランク", bonusFactorHeader, rankSelectionHeader, actionTriggerHeader, "操作結果", "操作日時", currentStatusHeader, sortPriorityHeader];
const terminationHeaders = ["社員ID", "表示名", "DiscordユーザーID", "最終ランク", "解雇日", "手続き完了", "対応署員", "完了日", "名簿削除予定日", "名簿削除状況", "備考"];
const terminationSheetName = "解雇者管理";
const employeeSheetId = 1100459512;
const employeeSheetName = "署員一覧";
const retentionPeriodMs = 7 * 24 * 60 * 60 * 1000;
const bonusDistributionSheetName = "ボーナス配布";
const bonusDistributionSheetId = 1863429017;
const bonusRoundSettingsSheetName = "ボーナス回設定";
let displayedBonusRound = "";
const recruitmentSettingsSheetName = "募集設定";
const applicationSheetName = "応募者管理";
const applicationSheetId = 2081134610;
let displayedApplicationRound = "";
let lastRecruitmentPollAt = 0;
const recruitmentPollIntervalMs = Math.max(Number(process.env.RECRUITMENT_POLL_INTERVAL_MS || 60000), 30000);
const interviewSettingsSheetName = "面接設定";
const interviewQuestionsSheetName = "面接質問設定";
const interviewManagementSheetName = "面接者管理";
const onboardingSheetName = "採用手続き管理";
const staffProfileSheetName = "署員個票";
const staffProfileSheetId = 2090134610;
const rankOperationConfigRange = `'${unifiedSettingsSheetName}'!A1440:B1445`;
const commandBoardConfigRange = `'${unifiedSettingsSheetName}'!A1448:B1455`;
const commandBoardCooldownMs = 5000;
const commandBoardCooldowns = new Map();
const commandBoardLocks = new Set();
const commandBoardChannels = new Map();
const onboardingHeaders = [
  "応募ID", "面接ID", "募集回", "受験者名", "Discordユーザー名", "DiscordユーザーID",
  "面接合格日時", "手続き完了", "対応署員", "完了日時", "採用状態", "ロール処理結果",
  "最終エラー", "呼出日時", "呼出者", "備考",
];
let lastInterviewPollAt = 0;
const interviewPollIntervalMs = Math.max(Number(process.env.INTERVIEW_POLL_INTERVAL_MS || 60000), 30000);
let lastRankValidationAt = 0;

async function ensureRankAndEmployeeValidation() {
  if (Date.now() - lastRankValidationAt < 10 * 60 * 1000) return;
  const employeeSheet = await readEmployees();
  const rankRows = unifiedSettingsSnapshot?.rank || (await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${unifiedSettingsSheetName}'!A10:F109`,
    valueRenderOption: "UNFORMATTED_VALUE",
  })).data.values || [];
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties(sheetId,title)",
  });
  const settingsSheetId = metadata.data.sheets?.find((item) => item.properties?.title === unifiedSettingsSheetName)?.properties?.sheetId;
  if (settingsSheetId === undefined || settingsSheetId === null) throw new Error("設定シートが見つかりません");
  const rankEnabledValues = Array.from({ length: 100 }, (_, index) => [isEnabledSetting(rankRows[index]?.[5])]);
  const rankCandidateFormulas = Array.from({ length: 100 }, (_, index) => [`=IF(OR($F${index + 10}=TRUE,$F${index + 10}=\"はい\"),$B${index + 10},\"\")`]);
  const enabledRankNames = rankRows
    .filter((row) => isEnabledSetting(row?.[5]) && String(row?.[1] || "").trim())
    .map((row) => String(row[1]).trim());
  const rankColumn = employeeSheet.headerMap.get(rankSelectionHeader);
  const triggerColumn = employeeSheet.headerMap.get(actionTriggerHeader);
  const requests = [
    {
      setDataValidation: {
        range: { sheetId: settingsSheetId, startRowIndex: 9, endRowIndex: 109, startColumnIndex: 5, endColumnIndex: 6 },
        rule: { condition: { type: "BOOLEAN" }, strict: true, showCustomUi: true },
      },
    },
    {
      updateDimensionProperties: {
        range: { sheetId: settingsSheetId, dimension: "COLUMNS", startIndex: 6, endIndex: 7 },
        properties: { hiddenByUser: true },
        fields: "hiddenByUser",
      },
    },
  ];
  if (rankColumn !== undefined) {
    requests.push({
      setDataValidation: {
        range: { sheetId: employeeSheetId, startRowIndex: 2, endRowIndex: 1000, startColumnIndex: rankColumn, endColumnIndex: rankColumn + 1 },
        rule: { condition: { type: "ONE_OF_LIST", values: enabledRankNames.map((value) => ({ userEnteredValue: value })) }, strict: false, showCustomUi: true },
      },
    });
  }
  if (triggerColumn !== undefined) {
    requests.push({
      setDataValidation: {
        range: { sheetId: employeeSheetId, startRowIndex: 2, endRowIndex: 1000, startColumnIndex: triggerColumn, endColumnIndex: triggerColumn + 1 },
        rule: { condition: { type: "BOOLEAN" }, strict: true, showCustomUi: true },
      },
    });
  }
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: [
        { range: `'${unifiedSettingsSheetName}'!F10:F109`, values: rankEnabledValues },
        { range: `'${unifiedSettingsSheetName}'!G9`, values: [["有効ランク候補（自動）"]] },
        { range: `'${unifiedSettingsSheetName}'!G10:G109`, values: rankCandidateFormulas },
      ],
    },
  });
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  lastRankValidationAt = Date.now();
  console.log("ランク設定の有効チェックボックスと署員一覧のランクプルダウンを更新しました");
}

async function readEmployees() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${employeeSheetName}'!A2:ZZ1000`,
  });
  const values = response.data.values || [];
  let headers = values[0] || [];
  const missingHeaders = botHeaders.filter((header) => {
    if (header === actionTriggerHeader) {
      return !headers.includes(actionTriggerHeader) && !headers.includes(legacyActionTriggerHeader);
    }
    if (header === bonusFactorHeader) {
      return !headers.includes(bonusFactorHeader) && !headers.includes(legacyBonusFactorHeader);
    }
    return !headers.includes(header);
  });
  if (missingHeaders.length) {
    const start = Math.max(headers.length, 1);
    const end = start + missingHeaders.length - 1;
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties(sheetId,title,gridProperties.columnCount)",
    });
    const sheet = metadata.data.sheets?.find((item) => item.properties?.title === employeeSheetName);
    const currentColumnCount = Number(sheet?.properties?.gridProperties?.columnCount || 0);
    const requiredColumnCount = end + 1;
    if (currentColumnCount < requiredColumnCount) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            insertDimension: {
              range: {
                sheetId: employeeSheetId,
                dimension: "COLUMNS",
                startIndex: currentColumnCount,
                endIndex: requiredColumnCount,
              },
              inheritFromBefore: true,
            },
          }],
        },
      });
    }
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${employeeSheetName}'!${columnLetter(start)}2:${columnLetter(end)}2`,
      valueInputOption: "RAW",
      requestBody: { values: [missingHeaders] },
    });
    headers = [...headers, ...missingHeaders];
    console.log(`見出しを自動追加: ${missingHeaders.join(", ")}`);
  }
  const headerMap = new Map(headers.map((header, index) => [String(header), index]));
  if (!headerMap.has(actionTriggerHeader) && headerMap.has(legacyActionTriggerHeader)) {
    headerMap.set(actionTriggerHeader, headerMap.get(legacyActionTriggerHeader));
  }
  if (!headerMap.has(bonusFactorHeader) && headerMap.has(legacyBonusFactorHeader)) {
    headerMap.set(bonusFactorHeader, headerMap.get(legacyBonusFactorHeader));
  }
  return {
    headers,
    headerMap,
    rows: values.slice(1),
  };
}

function cellRange(employeeSheet, header, rowNumber) {
  const index = employeeSheet.headerMap.get(header);
  if (index === undefined) throw new Error(`${employeeSheetName}シートに「${header}」列がありません。`);
  const col = columnLetter(index);
  return `'${employeeSheetName}'!${col}${rowNumber}`;
}

function rowUpdate(employeeSheet, rowNumber, fields) {
  return Object.entries(fields).map(([header, value]) => ({
    range: cellRange(employeeSheet, header, rowNumber),
    values: [[value]],
  }));
}

function ref(employeeSheet, header, rowNumber) {
  return `${columnLetter(employeeSheet.headerMap.get(header))}${rowNumber}`;
}

function employeeFormulas(employeeSheet, rowNumber) {
  const id = ref(employeeSheet, "社員ID", rowNumber);
  const appliedRank = ref(employeeSheet, "適用ランク", rowNumber);
  return {
    [bonusFactorHeader]: `=IF(${id}="","",IF('設定'!$K$3="v2:READY",IFNA(XLOOKUP(${appliedRank},'設定'!$B$10:$B$109,'設定'!$E$10:$E$109),0),IFNA(XLOOKUP(${appliedRank},'ランク設定'!$B$3:$B$1000,'ランク設定'!$E$3:$E$1000),0)))`,
  };
}

function employeeId(discordId) {
  return `DC-${discordId}`;
}

async function readTerminations() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${terminationSheetName}'!A2:K500`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const values = response.data.values || [];
  const headers = values[0] || terminationHeaders;
  return {
    headers,
    headerMap: new Map(headers.map((header, index) => [String(header), index])),
    rows: values.slice(1),
  };
}

function terminationCellRange(terminationSheet, header, rowNumber) {
  const index = terminationSheet.headerMap.get(header);
  if (index === undefined) throw new Error(`${terminationSheetName}シートに「${header}」列がありません。`);
  const col = columnLetter(index);
  return `'${terminationSheetName}'!${col}${rowNumber}`;
}

function terminationRowUpdate(terminationSheet, rowNumber, fields) {
  return Object.entries(fields).map(([header, value]) => ({
    range: terminationCellRange(terminationSheet, header, rowNumber),
    values: [[value]],
  }));
}

function discordIdCell(discordId) {
  return `'${discordId}`;
}

function normalizedDiscordId(row, employeeSheet) {
  const employeeIndex = employeeSheet.headerMap.get("社員ID");
  return String(row[employeeIndex] || "").replace(/^DC-/, "").trim();
}

function isChecked(value) {
  return value === true || String(value || "").toUpperCase() === "TRUE";
}

function parseSheetDate(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return new Date((value - 25569) * 24 * 60 * 60 * 1000 - 9 * 60 * 60 * 1000);
  }
  const text = String(value || "").trim();
  if (!text) return null;
  const localMatch = text.match(/^(\d{4})\/(\d{2})\/(\d{2})\s+(\d{2}):(\d{2})(?::(\d{2}))?$/);
  const parsed = localMatch
    ? new Date(`${localMatch[1]}-${localMatch[2]}-${localMatch[3]}T${localMatch[4]}:${localMatch[5]}:${localMatch[6] || "00"}+09:00`)
    : new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function sheetDateTime(date) {
  const parts = new Intl.DateTimeFormat("ja-JP-u-ca-gregory", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}/${values.month}/${values.day} ${values.hour}:${values.minute}:${values.second}`;
}

async function upsertTerminationRecord(member, finalRank = "") {
  const terminationSheet = await readTerminations();
  const employeeIdColumn = terminationSheet.headerMap.get("社員ID");
  const discordIdColumn = terminationSheet.headerMap.get("DiscordユーザーID");
  const finalRankColumn = terminationSheet.headerMap.get("最終ランク");
  const index = terminationSheet.rows.findIndex((row) => {
    const discordId = String(row[discordIdColumn] || "").replace(/^'/, "").trim();
    return discordId === member.id || String(row[employeeIdColumn] || "").trim() === employeeId(member.id);
  });
  const emptyIndex = terminationSheet.rows.findIndex((row) => !String(row[employeeIdColumn] || "").trim());
  const rowNumber = index >= 0 ? index + 3 : emptyIndex >= 0 ? emptyIndex + 3 : terminationSheet.rows.length + 3;
  const existingRank = index >= 0 ? String(terminationSheet.rows[index][finalRankColumn] || "").trim() : "";
  const fields = {
    "社員ID": employeeId(member.id),
    "表示名": member.displayName,
    "DiscordユーザーID": discordIdCell(member.id),
  };
  if (finalRank && !existingRank) fields["最終ランク"] = finalRank;
  if (index < 0) {
    fields["解雇日"] = sheetDateTime(new Date());
    fields["手続き完了"] = false;
  }
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: terminationRowUpdate(terminationSheet, rowNumber, fields),
    },
  });
  return rowNumber;
}

async function applyEmployeeUpdates(data, context) {
  if (context?.pendingData) {
    context.pendingData.push(...data);
    return;
  }
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: "USER_ENTERED", data },
  });
}

async function sortEmployees(employeeSheet) {
  const priorityColumn = employeeSheet.headerMap.get(sortPriorityHeader);
  const nameColumn = employeeSheet.headerMap.get("表示名");
  if (priorityColumn === undefined || nameColumn === undefined) {
    throw new Error(`${employeeSheetName}シートの並び替え用見出しが見つかりません`);
  }
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          updateDimensionProperties: {
            range: {
              sheetId: employeeSheetId,
              dimension: "COLUMNS",
              startIndex: priorityColumn,
              endIndex: priorityColumn + 1,
            },
            properties: { hiddenByUser: true },
            fields: "hiddenByUser",
          },
        },
        {
          sortRange: {
            range: {
              sheetId: employeeSheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: 0,
              endColumnIndex: Math.max(employeeSheet.headers.length, botHeaders.length),
            },
            sortSpecs: [
              { dimensionIndex: priorityColumn, sortOrder: "ASCENDING" },
              { dimensionIndex: nameColumn, sortOrder: "ASCENDING" },
            ],
          },
        },
      ],
    },
  });
}

async function clearEmployeeRow(employeeSheet, rowNumber, context = null) {
  const range = `'${employeeSheetName}'!A${rowNumber}:ZZ${rowNumber}`;
  if (context?.pendingClearRanges) {
    context.pendingClearRanges.push(range);
    return;
  }
  await sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    requestBody: { ranges: [range] },
  });
  await sortEmployees(employeeSheet);
}

function logMemberSync(context, message) {
  if (!context?.silent) console.log(message);
}

async function syncMember(member, context = null) {
  if (member.guild.id !== guildId || member.user.bot) return;
  const [rankMap, employeeSheet] = context
    ? [context.rankMap, context.employeeSheet]
    : await Promise.all([readRankMap(), readEmployees()]);
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const index = employeeSheet.rows.findIndex((row) => String(row[idColumn] || "") === employeeId(member.id));
  const emptyIndex = employeeSheet.rows.findIndex((row) => !row[idColumn]);
  const targetRow = index >= 0 ? index + 3 : emptyIndex >= 0 ? emptyIndex + 3 : employeeSheet.rows.length + 3;
  let hasRosterRole = [...rosterRoleIds].some((id) => member.roles.cache.has(id));
  const hasExcludeRole = [...excludeRoleIds].some((id) => member.roles.cache.has(id));
  const dismissal = dismissedRank(rankMap);
  const hasDismissedRole = Boolean(dismissal && member.roles.cache.has(dismissal.roleId));
  const assessed = assessMember(member, rankMap);
  const hasAnyRole = member.roles.cache.some((role) => role.id !== member.guild.id);

  if (!hasAnyRole) {
    if (index >= 0) await clearEmployeeRow(employeeSheet, targetRow, context);
    lastSynchronizedRanks.delete(member.id);
    logMemberSync(context, `ロールなしのため名簿削除: ${member.displayName}`);
    return;
  }

  if (hasDismissedRole) {
    if (hasRosterRole) {
      const removedRoles = await removeRosterRoles(
        member.guild,
        member,
        "退職者ロールを検知したため名簿対象ロールを自動解除",
      );
      hasRosterRole = false;
      logMemberSync(context, `退職者の名簿対象ロールを自動解除: ${member.displayName} (${removedRoles.join(", ")})`);
    }
    const appliedRankColumn = employeeSheet.headerMap.get("適用ランク");
    const previousRank = assessed.rankName || (index >= 0
      ? String(employeeSheet.rows[index][appliedRankColumn] || "").trim()
      : "");
    await upsertTerminationRecord(member, previousRank);
    if (index >= 0) {
      // 解雇者は署員一覧には残さず、解雇者管理だけを正とする。
      // 解除予定日まで解雇者管理に保持し、名簿側は検知時点で即時クリアする。
      await clearEmployeeRow(employeeSheet, targetRow, context);
    }
    lastSynchronizedRanks.set(member.id, "解雇");
    logMemberSync(context, `解雇者を署員一覧から除外: ${member.displayName}${previousRank ? ` (最終ランク: ${previousRank})` : ""}`);
    return;
  }

  if (!hasRosterRole && assessed.rankName && index >= 0 && !hasExcludeRole) {
    await ensureRosterRole(member.guild, member, "階級ロールが残っているためPolice Officerを自動復元");
    hasRosterRole = true;
    logMemberSync(context, `Police Officerロール自動復元: ${member.displayName}`);
  }

  const eligible = hasRosterRole && !hasExcludeRole;

  if (!eligible) {
    if (index < 0) return;
    await clearEmployeeRow(employeeSheet, targetRow, context);
    lastSynchronizedRanks.delete(member.id);
    logMemberSync(context, `名簿対象外: ${member.displayName}`);
    return;
  }

  const appliedRank = assessed.rankName || "？？？？";

  if (index < 0) {
    await applyEmployeeUpdates(
      rowUpdate(employeeSheet, targetRow, { "社員ID": employeeId(member.id), "表示名": member.displayName, "Discordロール": assessed.roleNames, "適用ランク": appliedRank, [sortPriorityHeader]: assessed.priority, ...employeeFormulas(employeeSheet, targetRow) }),
      context,
    );
    const rowIndex = targetRow - 3;
    while (employeeSheet.rows.length <= rowIndex) employeeSheet.rows.push([]);
    employeeSheet.rows[rowIndex][idColumn] = employeeId(member.id);
  } else {
    await applyEmployeeUpdates(
      rowUpdate(employeeSheet, targetRow, { "表示名": member.displayName, "Discordロール": assessed.roleNames, "適用ランク": appliedRank, [sortPriorityHeader]: assessed.priority }),
      context,
    );
  }
  lastSynchronizedRanks.set(member.id, appliedRank);
  logMemberSync(context, `同期完了: ${member.displayName} → ${appliedRank}`);
}

async function consolidateEmployeeDuplicates() {
  const employeeSheet = await readEmployees();
  const groups = new Map();
  employeeSheet.rows.forEach((row, index) => {
    const discordId = normalizedDiscordId(row, employeeSheet);
    if (!discordId) return;
    if (!groups.has(discordId)) groups.set(discordId, []);
    groups.get(discordId).push(index);
  });

  const duplicateGroups = [...groups.entries()].filter(([, indexes]) => indexes.length > 1);
  if (!duplicateGroups.length) return 0;

  const systemHeaders = new Set([
    "社員ID", "表示名", "Discordロール", "適用ランク",
    "基本ボーナス", bonusFactorHeader, legacyBonusFactorHeader, "調整額", "見込ボーナス", rankSelectionHeader, actionTriggerHeader, "操作結果", "操作日時", currentStatusHeader, sortPriorityHeader,
  ]);
  const mergeHeaders = employeeSheet.headers.filter((header) => header && !systemHeaders.has(String(header)));
  const mergeData = [];
  const duplicateIndexes = [];

  for (const [discordId, indexes] of duplicateGroups) {
    const canonicalIndex = indexes[0];
    const canonicalRow = employeeSheet.rows[canonicalIndex];
    const canonicalRowNumber = canonicalIndex + 3;
    const merged = {};
    for (const header of mergeHeaders) {
      const column = employeeSheet.headerMap.get(header);
      if (column === undefined) continue;
      const value = indexes
        .map((index) => employeeSheet.rows[index][column])
        .find((candidate) => candidate !== undefined && candidate !== null && String(candidate).trim() !== "");
      if (value !== undefined && String(canonicalRow[column] || "").trim() === "") merged[header] = value;
    }
    mergeData.push(...rowUpdate(employeeSheet, canonicalRowNumber, merged));
    duplicateIndexes.push(...indexes.slice(1));
    console.log(`重複統合予定: ${discordId} (${indexes.length}件)`);
  }

  if (mergeData.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data: mergeData },
    });
  }

  const uniqueDuplicateIndexes = [...new Set(duplicateIndexes)];
  await sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    requestBody: {
      ranges: uniqueDuplicateIndexes.map((index) => `'${employeeSheetName}'!A${index + 3}:ZZ${index + 3}`),
    },
  });
  await sortEmployees(employeeSheet);
  console.log(`重複統合完了: ${uniqueDuplicateIndexes.length}行を統合`);
  return uniqueDuplicateIndexes.length;
}

async function writeActionResult(employeeSheet, rowNumber, fields) {
  const row = employeeSheet.rows[rowNumber - 3] || [];
  const changedFields = Object.fromEntries(Object.entries(fields).filter(([header, value]) => {
    const column = employeeSheet.headerMap.get(header);
    if (column === undefined) return false;
    const previous = row[column];
    // Action/status cells are strings or booleans. Avoid a Sheets write when
    // the value already matches; this function runs for every employee every
    // polling cycle.
    if (typeof value === "boolean") return isChecked(previous) !== value;
    return String(previous ?? "").trim() !== String(value ?? "").trim();
  }));
  if (!Object.keys(changedFields).length) return false;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: rowUpdate(employeeSheet, rowNumber, changedFields),
    },
  });
  for (const [header, value] of Object.entries(changedFields)) {
    const column = employeeSheet.headerMap.get(header);
    if (column !== undefined) row[column] = value;
  }
  return true;
}

async function processSheetActions() {
  if (!client.isReady()) return;
  const [guild, rankMap, employeeSheet] = await Promise.all([
    client.guilds.fetch(guildId),
    readRankMap(),
    readEmployees(),
  ]);
  const targetColumn = employeeSheet.headerMap.get(rankSelectionHeader);
  const triggerColumn = employeeSheet.headerMap.get(actionTriggerHeader);
  const appliedRankColumn = employeeSheet.headerMap.get("適用ランク");
  let needsSort = false;

  for (let index = 0; index < employeeSheet.rows.length; index += 1) {
    const row = employeeSheet.rows[index];
    const rowNumber = index + 3;
    const discordId = normalizedDiscordId(row, employeeSheet);
    const triggered = isChecked(row[triggerColumn]);
    if (!discordId && triggered) {
      await writeActionResult(employeeSheet, rowNumber, {
        [actionTriggerHeader]: false,
        "操作結果": "エラー: 社員IDからDiscordユーザーIDを判別できません（IDが違います）",
        [currentStatusHeader]: "要確認: Discord IDが違います",
        "操作日時": new Date().toISOString(),
      });
      continue;
    }
    if (!discordId) {
      await writeActionResult(employeeSheet, rowNumber, { [currentStatusHeader]: "対象外: Discord ID未登録" });
      continue;
    }

    try {
      const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId);
      const dismissal = dismissedRank(rankMap);
      const discordRank = dismissal && member.roles.cache.has(dismissal.roleId)
        ? "解雇"
        : assessMember(member, rankMap).rankName || "？？？？";

      if (triggered) {
        const selectedRank = String(row[targetColumn] || "").trim();
        await writeActionResult(employeeSheet, rowNumber, { "操作結果": `処理中: ${selectedRank || "未選択"}`, [currentStatusHeader]: "処理中" });
        if (!selectedRank) throw new Error("変更後ランクを選択してください");
        const result = await applySelectedRank(guild, member, rankMap, selectedRank, "Google Sheetsの統合操作");
        await syncMember(await guild.members.fetch(discordId));
        needsSort = true;
        await writeActionResult(employeeSheet, rowNumber, {
          [rankSelectionHeader]: "",
          [actionTriggerHeader]: false,
          "操作結果": `完了: ${result.transition}`,
          [currentStatusHeader]: `在籍中: ${result.transition.split(" → ").at(-1) || "同期済み"}`,
          "操作日時": new Date().toISOString(),
        });
        console.log(`シート操作完了: ${member.displayName} ${result.transition}`);
        continue;
      }

      if (discordRank === "解雇") {
        lastSynchronizedRanks.set(discordId, discordRank);
        await writeActionResult(employeeSheet, rowNumber, { [currentStatusHeader]: "解雇者管理へ移管済み" });
        continue;
      }

      const sheetRank = String(row[appliedRankColumn] || "").trim();
      if (!sheetRank || sheetRank === discordRank) {
        lastSynchronizedRanks.set(discordId, discordRank);
        await writeActionResult(employeeSheet, rowNumber, { [currentStatusHeader]: `在籍中: ${discordRank}` });
        continue;
      }

      const lastRank = lastSynchronizedRanks.get(discordId);
      if (lastRank === discordRank) {
        const result = await applySelectedRank(guild, member, rankMap, sheetRank, "Google Sheetsの適用ランク直接編集");
        await syncMember(await guild.members.fetch(discordId));
        needsSort = true;
        await writeActionResult(employeeSheet, rowNumber, {
          "操作結果": `完了: スプシ → Discord (${result.transition})`,
          [currentStatusHeader]: `在籍中: ${sheetRank}`,
          "操作日時": new Date().toISOString(),
        });
        console.log(`双方向同期: スプシ → Discord ${member.displayName} ${result.transition}`);
      } else {
        await syncMember(member);
        needsSort = true;
        await writeActionResult(employeeSheet, rowNumber, {
          "操作結果": `完了: Discord → スプシ (${discordRank})`,
          [currentStatusHeader]: `在籍中: ${discordRank}`,
          "操作日時": new Date().toISOString(),
        });
        console.log(`双方向同期: Discord → スプシ ${member.displayName} ${discordRank}`);
      }
    } catch (error) {
      const message = formatDiscordIdError(error, discordId);
      await writeActionResult(employeeSheet, rowNumber, {
        [actionTriggerHeader]: false,
        "操作結果": `エラー: ${message}`,
        [currentStatusHeader]: /Discord IDが違います/.test(message) ? "要確認: Discord IDが違います" : `処理失敗: ${message}`,
        "操作日時": new Date().toISOString(),
      });
      console.error(`シート操作失敗: ${discordId}`, message);
    }
  }
  if (needsSort) await sortEmployees(await readEmployees());
}

async function applyBonusRoundFilter(roundName) {
  if (!roundName || displayedBonusRound === roundName) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        setBasicFilter: {
          filter: {
            range: {
              sheetId: bonusDistributionSheetId,
              startRowIndex: 10,
              endRowIndex: 1000,
              startColumnIndex: 0,
              endColumnIndex: 12,
            },
            criteria: {
              11: {
                condition: {
                  type: "TEXT_EQ",
                  values: [{ userEnteredValue: roundName }],
                },
              },
            },
          },
        },
      }],
    },
  });
  displayedBonusRound = roundName;
}

async function writeBonusTop(data) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: Object.entries(data).map(([cell, value]) => ({
        range: `'${bonusDistributionSheetName}'!${cell}`,
        values: [[value]],
      })),
    },
  });
}

async function saveBonusRoundSetting(roundName, poolAmount, loadedAt) {
  const unified = await unifiedRange("bonus");
  const response = unified || await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${bonusRoundSettingsSheetName}'!A2:C100`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  // values.get omits trailing empty rows.  The unified block is a fixed
  // 100-row area (A114:C213), so pad the snapshot before looking for the
  // first available row.  Without this, a valid new round was incorrectly
  // treated as a full 100-row block and produced the old upper-limit error.
  const rows = [...(unified?.rows || response.data.values || [])];
  const maxRows = unified ? 100 : 99;
  while (rows.length < maxRows) rows.push([]);
  let index = rows.findIndex((row) => String(row[0] || "").trim() === roundName);
  if (index < 0) index = rows.findIndex((row) => !String(row[0] || "").trim());
  if (index < 0 || index >= maxRows) throw new Error("ボーナス回設定が上限に達しています");
  const rowNumber = index + (unified?.firstRow || 2);
  const sheetName = unified ? unifiedSettingsSheetName : bonusRoundSettingsSheetName;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetName}'!A${rowNumber}:C${rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[roundName, poolAmount, sheetDateTime(loadedAt)]] },
  });
}

function bonusLedgerRow(rowNumber, roundName, employee, loadedAt, existingRow = null) {
  const basicPay = `=IF(B${rowNumber}="","",IF(SUMIF($L$12:$L$1000,$L${rowNumber},$E$12:$E$1000)=0,0,ROUNDDOWN(IFNA(XLOOKUP($L${rowNumber},'設定'!$A$114:$A$213,'設定'!$B$114:$B$213),0)*E${rowNumber}/SUMIF($L$12:$L$1000,$L${rowNumber},$E$12:$E$1000),-7)))`;
  return [
    sheetDateTime(loadedAt).slice(0, 10),
    employee.employeeId,
    employee.name,
    employee.rank,
    employee.factor,
    basicPay,
    Number(existingRow?.[6]) || 0,
    `=IF(B${rowNumber}="","",F${rowNumber}+G${rowNumber})`,
    String(existingRow?.[8] || "").trim(),
    isChecked(existingRow?.[9]),
    isChecked(existingRow?.[10]),
    roundName,
  ];
}

async function processBonusDistribution() {
  const topResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${bonusDistributionSheetName}'!A4:R4`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const top = topResponse.data.values?.[0] || [];
  const roundName = String(top[1] || "").trim();
  if (roundName) await applyBonusRoundFilter(roundName);
  if (!isChecked(top[10])) return;

  try {
    await writeBonusTop({ N4: "処理中: 署員一覧を読込中" });
    if (!roundName) throw new Error("表示回を選択してください");
    const poolAmount = Number(top[7]);
    if (!Number.isFinite(poolAmount) || poolAmount <= 0) {
      throw new Error("新規プール入力に0より大きい金額を入力してください");
    }

    const [employeeSheet, ledgerResponse] = await Promise.all([
      readEmployees(),
      sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${bonusDistributionSheetName}'!A12:L1000`,
        valueRenderOption: "UNFORMATTED_VALUE",
      }),
    ]);

    const employeeIdColumn = employeeSheet.headerMap.get("社員ID");
    const employeeNameColumn = employeeSheet.headerMap.get("表示名");
    const employeeRolesColumn = employeeSheet.headerMap.get("Discordロール");
    const employeeRankColumn = employeeSheet.headerMap.get("適用ランク");
    const employeeFactorColumn = employeeSheet.headerMap.get(bonusFactorHeader);
    const employees = employeeSheet.rows
      .map((row) => ({
        employeeId: String(row[employeeIdColumn] || "").trim(),
        name: String(row[employeeNameColumn] || "").trim(),
        roles: String(row[employeeRolesColumn] || "").trim(),
        rank: String(row[employeeRankColumn] || "").trim() || "？？？？",
        factor: Number(row[employeeFactorColumn]) || 0,
      }))
      .filter((employee) => employee.employeeId && employee.name
        && employee.rank !== "解雇者" && !employee.roles.includes("解雇者"));
    if (!employees.length) throw new Error("署員一覧に読込対象がいません");

    const ledgerRows = ledgerResponse.data.values || [];
    const oldRoundRows = [];
    const emptyRows = [];
    const existingByEmployeeId = new Map();
    for (let index = 0; index < 989; index += 1) {
      const row = ledgerRows[index] || [];
      const rowNumber = index + 12;
      const rowRound = String(row[11] || "").trim();
      if (rowRound === roundName) {
        oldRoundRows.push(rowNumber);
        const existingEmployeeId = String(row[1] || "").trim();
        if (existingEmployeeId && !existingByEmployeeId.has(existingEmployeeId)) {
          existingByEmployeeId.set(existingEmployeeId, { rowNumber, row });
        }
      }
      else if (!String(row[1] || "").trim() && !rowRound) emptyRows.push(rowNumber);
    }
    const availableRows = [...oldRoundRows, ...emptyRows];
    const usedRows = new Set();
    const assignments = employees.map((employee) => {
      const existing = existingByEmployeeId.get(employee.employeeId);
      if (existing && !usedRows.has(existing.rowNumber)) {
        usedRows.add(existing.rowNumber);
        return { employee, rowNumber: existing.rowNumber, existingRow: existing.row };
      }
      const rowNumber = availableRows.find((candidate) => !usedRows.has(candidate));
      if (!rowNumber) return null;
      usedRows.add(rowNumber);
      return { employee, rowNumber, existingRow: null };
    });
    if (assignments.some((assignment) => !assignment)) {
      throw new Error("ボーナス配布の保存行が不足しています");
    }

    const loadedAt = new Date();
    await saveBonusRoundSetting(roundName, poolAmount, loadedAt);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: assignments.map(({ employee, rowNumber, existingRow }) => {
          return {
            range: `'${bonusDistributionSheetName}'!A${rowNumber}:L${rowNumber}`,
            values: [bonusLedgerRow(rowNumber, roundName, employee, loadedAt, existingRow)],
          };
        }),
      },
    });

    const unusedOldRows = oldRoundRows.filter((rowNumber) => !usedRows.has(rowNumber));
    if (unusedOldRows.length) {
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId,
        requestBody: {
          ranges: unusedOldRows.map((rowNumber) => `'${bonusDistributionSheetName}'!A${rowNumber}:L${rowNumber}`),
        },
      });
    }
    displayedBonusRound = "";
    await applyBonusRoundFilter(roundName);
    await writeBonusTop({
      H4: "",
      K4: false,
      N4: `完了: ${employees.length}名を${roundName}へ読込`,
    });
    console.log(`ボーナス配布読込完了: ${roundName} ${employees.length}名 プール=${poolAmount}`);
  } catch (error) {
    const message = error.response?.data?.error?.message || error.message || String(error);
    await writeBonusTop({ K4: false, N4: `エラー: ${message}` });
    console.error("ボーナス配布読込失敗:", message);
  }
}

function extractSpreadsheetId(value) {
  const text = String(value || "").trim();
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(text) ? text : "";
}

function normalizeFormHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s　]/g, "")
    .replace(/[？?！!（）()「」『』・。、,.]/g, "");
}

function formValue(headers, row, aliases, fallbackIndex = -1) {
  const normalizedAliases = aliases.map(normalizeFormHeader);
  const exact = headers.findIndex((header) => normalizedAliases.includes(normalizeFormHeader(header)));
  const partial = exact >= 0 ? exact : headers.findIndex((header) => {
    const normalized = normalizeFormHeader(header);
    return normalizedAliases.some((alias) => normalized.includes(alias) || alias.includes(normalized));
  });
  const index = partial >= 0 ? partial : fallbackIndex;
  return index >= 0 ? row[index] ?? "" : "";
}

function discordNumericId(value) {
  return String(value || "").replace(/^'/, "").match(/\d{17,20}/)?.[0] || "";
}

function normalizeDiscordUsername(value) {
  return String(value || "").replace(/^'/, "").replace(/^@/, "").trim().toLowerCase();
}

async function resolveApplicantDiscordUserId(value) {
  const numericId = discordNumericId(value);
  if (numericId) return numericId;
  const target = normalizeDiscordUsername(value);
  if (!target) return "";

  const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);
  const findMember = (members) => {
    const usernameMatch = members.find((member) => (
      member.user.username.toLowerCase() === target
      || member.user.tag.toLowerCase() === target
    ));
    if (usernameMatch) return usernameMatch;

    const displayMatches = members.filter((member) => [member.user.globalName, member.displayName]
      .filter(Boolean)
      .some((name) => String(name).trim().toLowerCase() === target));
    return displayMatches.size === 1 ? displayMatches.first() : null;
  };

  // A full member fetch uses Discord's guild-member request gateway and is
  // rate-limited when several applicants pass in the same sync. Reuse the
  // member cache populated by startup/previous lookups before requesting it.
  const cachedMatch = findMember(guild.members.cache);
  if (cachedMatch) return cachedMatch.id;
  const members = await guild.members.fetch();
  return findMember(members)?.id || "";
}

function recruitmentSettingFromRow(row, index, firstRow = 4, source = "legacy") {
  const requestedDuration = Number(row[5]);
  const requestedVoteLimit = Number(row[6]);
  return {
    rowNumber: index + firstRow,
    source,
    enabled: String(row[0] || "").trim() === "はい",
    roundName: String(row[1] || "").trim(),
    responseSpreadsheetUrl: String(row[2] || "").trim(),
    responseSheetName: "",
    pollChannelId: String(row[3] || "").trim(),
    pollMessage: String(row[4] || "").trim() || "応募内容を確認し、投票してください。",
    pollDurationHours: Number.isFinite(requestedDuration) && requestedDuration >= 1 && requestedDuration <= 168
      ? Math.floor(requestedDuration)
      : 24,
    pollVoteLimit: Number.isFinite(requestedVoteLimit) && requestedVoteLimit >= 1 && requestedVoteLimit <= 1000
      ? Math.floor(requestedVoteLimit)
      : 5,
    passChannelId: String(row[7] || "").trim(),
    passMessage: String(row[8] || "").trim() || "合格が決定しました。今後の案内をご確認ください。",
    manualTrigger: isChecked(row[11]),
    documentPassRoleId: String(row[13] || "").replace(/^'/, "").trim(),
  };
}

function sheetCellInputValue(cell) {
  const richLink = (cell?.chipRuns || [])
    .map((run) => run.chip?.richLinkProperties?.uri)
    .find(Boolean);
  if (richLink) return richLink;
  const value = cell?.userEnteredValue || {};
  if (value.stringValue !== undefined) return value.stringValue;
  if (value.numberValue !== undefined) return value.numberValue;
  if (value.boolValue !== undefined) return value.boolValue;
  return "";
}

async function readRecruitmentSettings() {
  const unified = await unifiedRange("recruitment");
  if (unified) {
    return unified.rows
      .map((row, index) => recruitmentSettingFromRow(row, index, unified.firstRow, unified.source))
      .filter((setting) => setting.roundName);
  }
  const response = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${recruitmentSettingsSheetName}'!A4:N100`],
    includeGridData: true,
    fields: "sheets.data.rowData.values(userEnteredValue,chipRuns)",
  });
  const rows = response.data.sheets?.[0]?.data?.[0]?.rowData || [];
  return rows
    .map((row) => (row.values || []).map(sheetCellInputValue))
    .map(recruitmentSettingFromRow)
    .filter((setting) => setting.roundName);
}

async function writeRecruitmentStatus(setting, status) {
  const key = `${setting.source || "legacy"}:${setting.roundName}`;
  const now = Date.now();
  const previous = recruitmentStatusCache.get(key);
  if (previous && previous.status === status && now - previous.at < 900000) return;
  const sheetName = setting.source === "unified" ? unifiedSettingsSheetName : recruitmentSettingsSheetName;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetName}'!J${setting.rowNumber}:L${setting.rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[status, sheetDateTime(new Date()), false]] },
  });
  recruitmentStatusCache.set(key, { status, at: now });
}

async function applyApplicationRoundFilter(roundName) {
  if (!roundName || displayedApplicationRound === roundName) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        setBasicFilter: {
          filter: {
            range: {
              sheetId: applicationSheetId,
              startRowIndex: 9,
              endRowIndex: 1000,
              startColumnIndex: 0,
              endColumnIndex: 27,
            },
            criteria: {
              24: {
                condition: {
                  type: "TEXT_EQ",
                  values: [{ userEnteredValue: roundName }],
                },
              },
            },
          },
        },
      }],
    },
  });
  displayedApplicationRound = roundName;
}

async function resolveResponseSheetName(responseSpreadsheetId, preferredName) {
  const response = await sheets.spreadsheets.get({
    spreadsheetId: responseSpreadsheetId,
    fields: "sheets.properties(title,index)",
  });
  const properties = (response.data.sheets || []).map((sheet) => sheet.properties);
  const preferred = properties.find((property) => property.title === preferredName);
  if (preferred) return preferred.title;
  const formResponses = properties.find((property) => /フォームの回答|form responses/i.test(property.title));
  if (formResponses) return formResponses.title;
  if (!properties.length) throw new Error("回答スプレッドシートにシートがありません");
  return properties.sort((a, b) => a.index - b.index)[0].title;
}

function quoteSheetName(sheetName) {
  return `'${String(sheetName).replace(/'/g, "''")}'`;
}

function responseApplication(headers, row) {
  return {
    submittedAt: formValue(headers, row, ["タイムスタンプ", "timestamp"], 0),
    discordId: String(formValue(headers, row, ["Discord ID", "DiscordユーザーID", "Discordユーザー名", "Discord username"])).replace(/^'/, "").trim(),
    name: String(formValue(headers, row, ["街での名前", "名前"])).trim(),
    age: String(formValue(headers, row, ["年齢"])).trim(),
    department: String(formValue(headers, row, ["希望の部署", "希望部署"])).trim(),
    availability: String(formValue(headers, row, ["1週間での起床率と出勤可能時間", "起床率", "出勤可能時間"])).trim(),
    experience: String(formValue(headers, row, ["PDの経験はありますか", "PD経験"])).trim(),
    reason: String(formValue(headers, row, ["志望理由"])).trim(),
    strength: String(formValue(headers, row, ["PD業務で得意なこと", "得意なこと"])).trim(),
    prosCons: String(formValue(headers, row, ["自分の長所と短所", "長所と短所"])).trim(),
    ideal: String(formValue(headers, row, ["あなたが希望する部署の理想のPDは", "理想のPD"])).trim(),
    whitelist: String(formValue(headers, row, ["ホワイトリスト申請の文言は", "ホワイトリスト"])).trim(),
    question: String(formValue(headers, row, ["質問等あれば記入して下さい", "質問"])).trim(),
  };
}

function shortHash(value) {
  let hash = 0;
  for (const character of String(value)) hash = ((hash << 5) - hash + character.charCodeAt(0)) | 0;
  return Math.abs(hash).toString(36).toUpperCase().padStart(6, "0").slice(-6);
}

function applicationId(roundName, sourceKey) {
  const roundNumber = String(roundName).match(/\d+/)?.[0]?.padStart(2, "0") || "00";
  return `APP-${roundNumber}-${shortHash(sourceKey)}`;
}

function applicationSheetRow(rowNumber, setting, application, sourceKey) {
  return [
    applicationId(setting.roundName, sourceKey),
    application.submittedAt,
    application.discordId,
    application.name,
    application.age,
    application.department,
    application.availability,
    application.experience,
    application.reason,
    application.strength,
    application.prosCons,
    application.ideal,
    application.whitelist,
    application.question,
    "PENDING",
    0,
    0,
    0,
    "投票待ち",
    "",
    "",
    "",
    "",
    "取込完了・投票作成待ち",
    setting.roundName,
    sourceKey,
    "",
    "",
    "",
    "",
  ];
}

function applicationFromSheetRow(row, rowNumber) {
  return {
    rowNumber,
    id: String(row[0] || "").trim(),
    submittedAt: row[1] || "",
    discordId: String(row[2] || "").replace(/^'/, "").trim(),
    name: String(row[3] || "").trim(),
    age: String(row[4] || "").trim(),
    department: String(row[5] || "").trim(),
    availability: String(row[6] || "").trim(),
    experience: String(row[7] || "").trim(),
    reason: String(row[8] || "").trim(),
    strength: String(row[9] || "").trim(),
    prosCons: String(row[10] || "").trim(),
    ideal: String(row[11] || "").trim(),
    whitelist: String(row[12] || "").trim(),
    question: String(row[13] || "").trim(),
    pollStatus: String(row[14] || "").trim(),
    passVotes: Number(row[15]) || 0,
    failVotes: Number(row[16]) || 0,
    totalVotes: Number(row[17]) || 0,
    verdict: String(row[18] || "").trim(),
    pollUrl: String(row[19] || "").trim(),
    pollEndsAt: row[20] || "",
    pollMessageId: String(row[21] || "").replace(/^'/, "").trim(),
    pollChannelId: String(row[22] || "").replace(/^'/, "").trim(),
    processResult: String(row[23] || "").trim(),
    roundName: String(row[24] || "").trim(),
    passAnnouncementMessageId: String(row[26] || "").replace(/^'/, "").trim(),
    resolvedDiscordId: String(row[27] || "").replace(/^'/, "").trim(),
    documentRoleStatus: String(row[28] || "").trim(),
    roleUpdatedAt: row[29] || "",
  };
}

function truncateDiscord(value, maxLength = 360) {
  const text = String(value || "").trim() || "（未記入）";
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function discordNonce(prefix, source) {
  const digest = createHash("sha256").update(`${prefix}:${source}`).digest("hex").slice(0, 16);
  return `${String(prefix || "n").slice(0, 8)}-${digest}`;
}

function applicationLink(rowNumber) {
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${applicationSheetId}&range=A${rowNumber}:AD${rowNumber}`;
}

function staffProfileLink() {
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=2090134610&range=B4`;
}

async function staffProfile(discordId) {
  const [employeeSheet, reportSummary] = await Promise.all([readEmployees(), readStaffReportSummary(discordId)]);
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const matchingRows = employeeSheet.rows.filter((item) => String(item[idColumn] || "").trim() === employeeId(discordId));
  if (!matchingRows.length) return null;
  if (matchingRows.length > 1) throw new Error(`社員ID ${employeeId(discordId)} が署員一覧で重複しています。管理者が一覧を確認してください。`);
  const [row] = matchingRows;
  const value = (header) => {
    const column = employeeSheet.headerMap.get(header);
    return column === undefined ? "" : row[column];
  };
  return {
    employeeId: String(value("社員ID") || ""),
    discordId,
    name: String(value("表示名") || ""),
    roles: String(value("Discordロール") || ""),
    rank: String(value("適用ランク") || "？？？？"),
    factor: value(bonusFactorHeader),
    report: reportSummary || String(value("報告欄") || ""),
    operationResult: String(value("操作結果") || ""),
    operationAt: value("操作日時"),
  };
}

function staffProfileEmbed(profile, member) {
  return new EmbedBuilder()
    .setColor(0x2563eb)
    .setTitle(`署員個票｜${truncateDiscord(profile.name || member.displayName, 220)}`)
    .setURL(staffProfileLink())
    .setThumbnail(member.displayAvatarURL({ size: 256 }))
    .addFields(
      { name: "社員ID", value: profile.employeeId, inline: true },
      { name: "DiscordユーザーID", value: profile.discordId, inline: true },
      { name: "現在の階級", value: profile.rank || "？？？？", inline: true },
      { name: "Discord表示名", value: member.displayName, inline: true },
      { name: "ユーザー名", value: member.user.username, inline: true },
      { name: "ランク係数", value: String(profile.factor ?? 0), inline: true },
      { name: "階級ロール", value: truncateDiscord(profile.roles || "設定なし", 1024) },
      { name: "報告欄", value: truncateDiscord(profile.report || "報告なし", 1024) },
      { name: "最終操作", value: truncateDiscord([profile.operationResult, profile.operationAt].filter(Boolean).join("｜") || "履歴なし", 1024) },
    )
    .setFooter({ text: "署員一覧の現在値を表示しています" })
    .setTimestamp();
}

async function readRankOperationSettings() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: rankOperationConfigRange,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const values = response.data.values || [];
  const map = new Map(values.map((row) => [String(row[0] || "").trim(), normalizedId(row[1])]));
  const reportChannelId = map.get("ランク報告チャンネルID") || "";
  const managementRoleId = map.get("ランク操作管理ロールID") || "";
  const promotionThreadId = map.get("昇格報告スレッドID") || "";
  const demotionThreadId = map.get("降格報告スレッドID") || "";
  const warningThreadId = map.get("警告報告スレッドID") || "";
  const reportThreadId = map.get("一般報告スレッドID") || "";
  return { reportChannelId, managementRoleId, promotionThreadId, demotionThreadId, warningThreadId, reportThreadId };
}

async function readCommandBoardSettings() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: commandBoardConfigRange,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const values = response.data.values || [];
  const map = new Map(values.map((row) => [String(row[0] || "").trim(), row[1]]));
  return {
    interviewChannelId: normalizedId(map.get("面接ボードチャンネルID")),
    interviewMessageId: normalizedId(map.get("面接ボードメッセージID")),
    interviewEnabled: isEnabledSetting(map.get("面接ボード有効")),
    rankChannelId: normalizedId(map.get("ランク管理ボードチャンネルID")),
    rankMessageId: normalizedId(map.get("ランク管理ボードメッセージID")),
    rankEnabled: isEnabledSetting(map.get("ランク管理ボード有効")),
  };
}

function commandBoardEmbed(kind) {
  const interview = kind === "interview";
  return new EmbedBuilder()
    .setColor(interview ? 0x2563eb : 0x15803d)
    .setTitle(interview ? "XPD管理｜面接ボード" : "XPD管理｜ランク管理ボード")
    .setDescription(`このチャンネルは${interview ? "面接" : "ランク管理"}専用です。権限のない機能は実行できません。\n連打防止のため、ボタンは5秒に1回までです。`)
    .addFields(...(interview
      ? [{ name: "面接（/mensetu）", value: "書類合格者を選択して面接を開始します。質問に回答後、投票へ進みます。", inline: false }]
      : [
        { name: "ランク操作（/rank）", value: "署員を選び、昇格・降格・警告・報告を実行します。備考は必須です。", inline: false },
        { name: "署員個票（/syoin）", value: "署員一覧から対象者を選択し、プロフィール・ランク変更履歴・報告欄を表示します。", inline: false },
      ]))
    .setFooter({ text: interview ? "詳細検索は /mensetu を使用してください" : "検索は /rank・/syoin を使用してください" })
    .setTimestamp();
}

function commandBoardComponents(kind) {
  return [new ActionRowBuilder().addComponents(...(kind === "interview"
    ? [new ButtonBuilder().setCustomId("command:mensetu").setLabel("面接を開始").setStyle(ButtonStyle.Primary)]
    : [
      new ButtonBuilder().setCustomId("command:rank").setLabel("ランク操作").setStyle(ButtonStyle.Success),
      new ButtonBuilder().setCustomId("command:syoin").setLabel("署員個票").setStyle(ButtonStyle.Secondary),
    ]))];
}

async function ensureCommandBoard(guild, kind, channelId, messageId, enabled) {
  const interview = kind === "interview";
  if (!enabled) {
    console.log(`${interview ? "面接" : "ランク管理"}ボード: 設定で無効です`);
    return;
  }
  if (!/^\d{17,20}$/.test(channelId)) {
    if (interview) {
      const interviewSettings = await readInterviewSettings();
      channelId = interviewSettings.find((setting) => setting.enabled && /^\d{17,20}$/.test(setting.commandChannelId))?.commandChannelId || "";
    } else {
      channelId = (await readRankOperationSettings()).reportChannelId;
    }
  }
  if (!/^\d{17,20}$/.test(channelId)) throw new Error(`${interview ? "面接" : "ランク管理"}ボードチャンネルIDが未設定です`);
  const channel = await guild.channels.fetch(channelId);
  if (!channel?.isTextBased?.() || channel.isThread?.()) throw new Error(`${interview ? "面接" : "ランク管理"}ボードチャンネルIDが無効です`);
  const payload = { embeds: [commandBoardEmbed(kind)], components: commandBoardComponents(kind) };
  let message = null;
  if (/^\d{17,20}$/.test(messageId)) message = await channel.messages.fetch(messageId).catch(() => null);
  if (message) await message.edit(payload);
  else {
    message = await channel.send(payload);
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${unifiedSettingsSheetName}'!${interview ? "B1451" : "B1454"}`,
      valueInputOption: "RAW",
      requestBody: { values: [[message.id]] },
    });
  }
  // Keep exactly one board for each configured channel. Older deployments could
  // have left duplicate bot-authored boards behind; remove only messages that
  // match this board title and never remove the configured/current message.
  const boardTitle = commandBoardEmbed(kind).data.title;
  const recent = await channel.messages.fetch({ limit: 50 }).catch(() => null);
  if (recent) {
    const duplicates = recent.filter((candidate) => (
      candidate.id !== message.id
      && candidate.author?.id === client.user?.id
      && candidate.embeds?.some((embed) => embed.title === boardTitle)
    ));
    for (const duplicate of duplicates.values()) {
      await duplicate.delete().catch((error) => console.warn(`${interview ? "面接" : "ランク管理"}ボード重複削除失敗 (${duplicate.id}):`, error.message));
    }
  }
  commandBoardChannels.set(interview ? "mensetu" : "rank", channel.id);
  if (!interview) commandBoardChannels.set("syoin", channel.id);
  console.log(`${interview ? "面接" : "ランク管理"}ボードを設置/更新: #${channel.name} (${message.id})`);
}

function commandBoardCooldownKey(interaction) {
  return interaction.user.id;
}

function commandBoardCooldownRemaining(interaction) {
  const key = commandBoardCooldownKey(interaction);
  const remaining = (commandBoardCooldowns.get(key) || 0) - Date.now();
  return remaining > 0 ? Math.ceil(remaining / 1000) : 0;
}

function consumeCommandBoardCooldown(interaction) {
  const key = commandBoardCooldownKey(interaction);
  const remaining = commandBoardCooldownRemaining(interaction);
  if (remaining > 0) return remaining;
  commandBoardCooldowns.set(key, Date.now() + commandBoardCooldownMs);
  setTimeout(() => commandBoardCooldowns.delete(key), commandBoardCooldownMs + 1000).unref?.();
  return 0;
}

async function validateRankReportDestinations(guild) {
  const settings = await readRankOperationSettings();
  if (!/^\d{17,20}$/.test(settings.reportChannelId)) throw new Error("ランク報告チャンネルIDが未設定です");
  const channel = await guild.channels.fetch(settings.reportChannelId).catch(() => null);
  if (!channel || channel.type !== ChannelType.GuildText) throw new Error("ランク報告チャンネルIDが無効です");
  const destinations = [
    ["昇格", settings.promotionThreadId],
    ["降格", settings.demotionThreadId],
    ["警告", settings.warningThreadId],
    ["報告", settings.reportThreadId],
  ];
  for (const [name, threadId] of destinations) {
    if (!/^\d{17,20}$/.test(threadId)) throw new Error(`${name}報告スレッドIDが未設定です`);
    const thread = await guild.channels.fetch(threadId).catch(() => null);
    if (!thread?.isThread?.() || thread.parentId !== channel.id) throw new Error(`${name}報告スレッドがチャンネル設定と一致しません`);
  }
  console.log(`ランク報告先を確認: #${channel.name} / ${destinations.map(([name]) => name).join("・")}`);
  return { channel };
}

function rankOperationPermission(interaction, settings) {
  if (interaction.memberPermissions?.has(PermissionFlagsBits.Administrator)) return true;
  return Boolean(settings.managementRoleId && interaction.member?.roles?.cache?.has(settings.managementRoleId));
}

function rankOperationType(rankMap, previousRank, nextRank) {
  const before = rankByName(rankMap, previousRank);
  const after = rankByName(rankMap, nextRank);
  if (!before || !after || before.priority === after.priority) return "ランク変更";
  return after.priority < before.priority ? "昇格" : "降格";
}

async function appendStaffReport(fields) {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${staffProfileSheetName}'!J:T`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [[
      fields.at,
      discordIdCell(fields.discordId),
      fields.name,
      fields.actorId,
      fields.actorName,
      fields.previousRank,
      fields.nextRank,
      fields.reason,
      fields.type,
      fields.threadUrl || "",
      fields.result,
    ]] },
  });
}

async function readStaffReportSummary(discordId) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${staffProfileSheetName}'!J2:T1000`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    return (response.data.values || [])
      .filter((row) => String(row[1] || "").replace(/^'/, "").trim() === discordId)
      .slice(-5)
      .reverse()
      .map((row) => `[${row[0] || "日時不明"}] ${row[8] || "ランク変更"}: ${row[5] || "？？？？"} → ${row[6] || "？？？？"} / ${row[7] || "備考なし"}`)
      .join("\n");
  } catch (error) {
    if (error.code === 400 || /Unable to parse range|not found/i.test(error.message || "")) return "";
    throw error;
  }
}

function applicationEmbed(application, title, color) {
  return new EmbedBuilder()
    .setColor(color)
    .setTitle(title)
    .setURL(applicationLink(application.rowNumber))
    .addFields(
      { name: "応募ID", value: truncateDiscord(application.id, 100), inline: true },
      { name: "Discordユーザー名", value: truncateDiscord(application.discordId, 100), inline: true },
      { name: "街での名前", value: truncateDiscord(application.name, 100), inline: true },
      { name: "年齢", value: truncateDiscord(application.age, 100), inline: true },
      { name: "希望部署", value: truncateDiscord(application.department, 100), inline: true },
      { name: "PD経験", value: truncateDiscord(application.experience, 100), inline: true },
      { name: "起床率・出勤可能時間", value: truncateDiscord(application.availability) },
      { name: "志望理由", value: truncateDiscord(application.reason) },
      { name: "PD業務で得意なこと", value: truncateDiscord(application.strength) },
      { name: "長所と短所", value: truncateDiscord(application.prosCons) },
      { name: "理想のPD", value: truncateDiscord(application.ideal) },
      { name: "ホワイトリスト文言", value: truncateDiscord(application.whitelist) },
      { name: "質問", value: truncateDiscord(application.question) },
    )
    .setFooter({ text: `${application.roundName}・応募管理シートへ移動できます` })
    .setTimestamp();
}

async function textChannel(channelId, label) {
  if (!/^\d{17,20}$/.test(channelId)) throw new Error(`${label}チャンネルIDが未設定です`);
  const channel = client.channels.cache.get(channelId) || await client.channels.fetch(channelId);
  if (!channel?.isTextBased() || typeof channel.send !== "function") {
    throw new Error(`${label}チャンネルへ送信できません: ${channelId}`);
  }
  return channel;
}

function pollVerdict(passVotes, totalVotes) {
  if (totalVotes <= 0) return "投票待ち";
  return passVotes * 2 >= totalVotes ? "合格" : "不合格";
}

function pollStateFromMessage(message, fetchedVotes = null) {
  const poll = message.poll;
  if (!poll) throw new Error("Discordメッセージに投票データがありません");
  const answers = [...poll.answers.values()];
  const passVotes = fetchedVotes?.passVotes
    ?? answers.find((answer) => answer.text === "合格")?.voteCount
    ?? 0;
  const failVotes = fetchedVotes?.failVotes
    ?? answers.find((answer) => answer.text === "不合格")?.voteCount
    ?? 0;
  const totalVotes = passVotes + failVotes;
  const expiresTimestamp = poll.expiresTimestamp || 0;
  const expired = expiresTimestamp > 0 && expiresTimestamp <= Date.now();
  const finalized = Boolean(poll.resultsFinalized) || expired;
  return {
    pollStatus: finalized ? "FINAL" : "PREVIEW",
    passVotes,
    failVotes,
    totalVotes,
    verdict: pollVerdict(passVotes, totalVotes),
    pollUrl: message.url,
    pollEndsAt: expiresTimestamp ? sheetDateTime(new Date(expiresTimestamp)) : "",
    pollMessageId: message.id,
    pollChannelId: message.channelId,
    processResult: finalized
      ? (expired ? "投票期限に到達して確定" : "投票結果を確定")
      : "投票結果を自動読込中",
  };
}

async function countPollAnswerVoters(answer) {
  if (!answer) return 0;
  let total = 0;
  let after = "";
  while (true) {
    const voters = await answer.fetchVoters({
      limit: 100,
      ...(after ? { after } : {}),
    });
    total += voters.size;
    const lastId = voters.lastKey();
    if (voters.size < 100 || !lastId || lastId === after) return total;
    after = lastId;
  }
}

async function fetchPollVoteCounts(message) {
  const answers = [...message.poll.answers.values()];
  const passAnswer = answers.find((answer) => answer.text === "合格");
  const failAnswer = answers.find((answer) => answer.text === "不合格");
  const [passVotes, failVotes] = await Promise.all([
    countPollAnswerVoters(passAnswer),
    countPollAnswerVoters(failAnswer),
  ]);
  return { passVotes, failVotes };
}

async function createApplicationPoll(setting, application) {
  const channel = await textChannel(setting.pollChannelId, "投票");
  const subject = truncateDiscord(application.name || application.discordId || application.id, 220);
  const message = await channel.send({
    content: setting.pollMessage,
    embeds: [applicationEmbed(application, `${setting.roundName} 応募審査`, 0x2563eb)],
    poll: {
      question: { text: `${subject} を合格としますか？` },
      answers: [
        { text: "合格", emoji: "✅" },
        { text: "不合格", emoji: "❌" },
      ],
      allowMultiselect: false,
      duration: setting.pollDurationHours,
    },
    allowedMentions: { parse: [] },
  });
  return pollStateFromMessage(message);
}

async function fetchApplicationPoll(application, setting) {
  const channel = await textChannel(application.pollChannelId, "投票");
  if (!channel.messages || typeof channel.messages.fetch !== "function") {
    throw new Error("投票メッセージを取得できないチャンネルです");
  }
  // Poll vote counts change without replacing the cached Message instance.
  // Always fetch from Discord so the sheet does not keep reading the 0-vote
  // snapshot that was cached when the bot created the poll.
  const message = await channel.messages.fetch({
    message: application.pollMessageId,
    force: true,
  });
  const fetchedVotes = await fetchPollVoteCounts(message);
  const state = pollStateFromMessage(message, fetchedVotes);
  if (state.pollStatus === "PREVIEW" && state.totalVotes >= setting.pollVoteLimit) {
    const endedMessage = await message.poll.end();
    return {
      ...pollStateFromMessage(endedMessage, fetchedVotes),
      pollStatus: "FINAL",
      processResult: `締切投票数${setting.pollVoteLimit}票に到達して確定`,
    };
  }
  return state;
}

async function sendApplicationPass(setting, application, { resend = false } = {}) {
  const channel = await textChannel(setting.passChannelId, "合格発表");
  const userId = await resolveApplicantDiscordUserId(application.discordId);
  const content = [setting.passMessage || "合格が決定しました。", userId ? `<@${userId}>` : ""]
    .filter(Boolean)
    .join("\n");
  const message = await channel.send({
    content,
    allowedMentions: { users: userId ? [userId] : [] },
    // A resend must get a fresh nonce; reusing the original nonce can be
    // deduplicated by Discord and appear as if nothing was sent.
    nonce: discordNonce("docpass", resend ? `${application.id}:${Date.now()}` : application.id),
    enforceNonce: true,
  });
  return { messageId: message.id, applicantMatched: Boolean(userId) };
}

function announcementRetryComponents(kind, id) {
  const prefix = kind === "interview" ? "iv" : "doc";
  return [new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`resend:${prefix}:${id}`)
      .setLabel("合格発表を再送信")
      .setStyle(ButtonStyle.Danger),
  )];
}

async function setAnnouncementRetryButton(channelId, messageId, kind, id, enabled) {
  if (!/^\d{17,20}$/.test(String(channelId || "")) || !/^\d{17,20}$/.test(String(messageId || ""))) return false;
  const channel = await textChannel(channelId, kind === "interview" ? "面接投票" : "投票");
  const message = await channel.messages.fetch({ message: messageId, force: true });
  await message.edit({ components: enabled ? announcementRetryComponents(kind, id) : [] });
  return true;
}

async function sanitizeApplicationPass(setting, application) {
  const channel = await textChannel(setting.passChannelId, "合格発表");
  if (!channel.messages || typeof channel.messages.fetch !== "function") return false;
  const message = await channel.messages.fetch({
    message: application.passAnnouncementMessageId,
    force: true,
  });
  if (message.author?.id !== client.user?.id) return false;

  const userId = await resolveApplicantDiscordUserId(application.discordId);
  const content = [setting.passMessage || "合格が決定しました。", userId ? `<@${userId}>` : ""]
    .filter(Boolean)
    .join("\n");
  if (message.content !== content || message.embeds.length > 0) {
    await message.edit({
      content,
      embeds: [],
      allowedMentions: { users: userId ? [userId] : [] },
    });
  }
  return Boolean(userId);
}

async function writeApplicationAnnouncementState(rowNumber, messageId, result) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: [
        { range: `'${applicationSheetName}'!X${rowNumber}`, values: [[result]] },
        { range: `'${applicationSheetName}'!AA${rowNumber}`, values: [[discordIdCell(messageId)]] },
      ],
    },
  });
}

async function updateApplicationRoleState(application, discordUserId, status) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: [
        { range: `'${applicationSheetName}'!AB${application.rowNumber}`, values: [[discordUserId ? discordIdCell(discordUserId) : ""]] },
        { range: `'${applicationSheetName}'!AC${application.rowNumber}`, values: [[status]] },
        { range: `'${applicationSheetName}'!AD${application.rowNumber}`, values: [[sheetDateTime(new Date())]] },
      ],
    },
  });
  application.resolvedDiscordId = discordUserId || "";
  application.documentRoleStatus = status;
}

async function resolveApplicationMember(guild, application) {
  const userId = application.resolvedDiscordId || await resolveApplicantDiscordUserId(application.discordId);
  if (!userId) throw new Error(`Discordユーザーを一意に特定できません: ${application.discordId || application.id}`);
  const member = guild.members.cache.get(userId) || await guild.members.fetch(userId);
  return { userId, member };
}

function sameApplicant(application, other, resolvedUserId) {
  if (other.id === application.id) return false;
  if (resolvedUserId && other.resolvedDiscordId === resolvedUserId) return true;
  const left = normalizeDiscordUsername(application.discordId);
  const right = normalizeDiscordUsername(other.discordId);
  return Boolean(left && right && left === right);
}

function isActiveStep1Pass(application) {
  return application.pollStatus === "FINAL"
    && application.verdict === "合格"
    && !application.documentRoleStatus.startsWith("面接結果反映済")
    && !application.documentRoleStatus.startsWith("採用完了");
}

async function reconcileStep1DocumentRole(guild, setting, application, allApplications = []) {
  if (application.pollStatus !== "FINAL" || !["合格", "不合格"].includes(application.verdict)) return false;
  const documentRoleId = configuredRoleId(guild, setting.documentPassRoleId, ["書類合格者", "書類合格"]);
  if (!documentRoleId) {
    if (application.verdict === "合格" && !application.documentRoleStatus) {
      await updateApplicationRoleState(application, application.resolvedDiscordId, "書類合格ロール設定待ち");
    }
    return false;
  }
  if (application.documentRoleStatus.startsWith("面接結果反映済") || application.documentRoleStatus.startsWith("採用完了")) return false;
  try {
    const { userId, member } = await resolveApplicationMember(guild, application);
    if (application.verdict === "合格") {
      const result = await reconcileMemberRoles(guild, member, { add: [documentRoleId] }, `STEP1書類合格: ${application.id}`);
      const status = "書類合格ロール付与済";
      if (application.documentRoleStatus !== status || application.resolvedDiscordId !== userId) {
        await updateApplicationRoleState(application, userId, status);
      }
      return result.added.length > 0;
    } else {
      const anotherPassExists = allApplications.some((other) => (
        isActiveStep1Pass(other) && sameApplicant(application, other, userId)
      ));
      if (anotherPassExists) {
        const status = "書類不合格・別の合格応募があるためロール維持";
        if (application.documentRoleStatus !== status || application.resolvedDiscordId !== userId) {
          await updateApplicationRoleState(application, userId, status);
        }
        return false;
      }
      const result = await reconcileMemberRoles(guild, member, { remove: [documentRoleId] }, `STEP1書類不合格: ${application.id}`);
      const status = "書類不合格・ロール解除済";
      if (application.documentRoleStatus !== status || application.resolvedDiscordId !== userId) {
        await updateApplicationRoleState(application, userId, status);
      }
      return result.removed.length > 0;
    }
  } catch (error) {
    await updateApplicationRoleState(application, application.resolvedDiscordId, `書類ロールエラー: ${error.message}`);
    return false;
  }
}

function applicationPollStateChanged(application, state) {
  return application.pollStatus !== state.pollStatus
    || application.passVotes !== state.passVotes
    || application.failVotes !== state.failVotes
    || application.totalVotes !== state.totalVotes
    || application.verdict !== state.verdict
    || application.pollUrl !== state.pollUrl
    || application.pollMessageId !== state.pollMessageId
    || application.pollChannelId !== state.pollChannelId;
}

async function writeApplicationPollState(rowNumber, state) {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${applicationSheetName}'!O${rowNumber}:X${rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        state.pollStatus,
        state.passVotes,
        state.failVotes,
        state.totalVotes,
        state.verdict,
        state.pollUrl,
        state.pollEndsAt,
        discordIdCell(state.pollMessageId),
        discordIdCell(state.pollChannelId),
        state.processResult,
      ]],
    },
  });
}

async function processRecruitmentApplications() {
  const now = Date.now();
  if (now - lastRecruitmentPollAt < recruitmentPollIntervalMs) return;
  lastRecruitmentPollAt = now;
  const [topResponse, settings, applicationResponse] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${applicationSheetName}'!A4:B4`,
      valueRenderOption: "UNFORMATTED_VALUE",
    }),
    readRecruitmentSettings(),
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${applicationSheetName}'!A11:AD1000`,
      valueRenderOption: "UNFORMATTED_VALUE",
    }),
  ]);
  const selectedRound = String(topResponse.data.values?.[0]?.[1] || "").trim();
  if (selectedRound) await applyApplicationRoundFilter(selectedRound);
  const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);

  const rows = applicationResponse.data.values || [];
  const sourceKeys = new Set(rows.map((row) => String(row[25] || "").trim()).filter(Boolean));
  const occupiedRows = new Set(rows.map((row, index) => String(row[0] || "").trim() ? index : -1).filter((index) => index >= 0));

  for (const setting of settings) {
    if (!setting.enabled && !setting.manualTrigger) continue;
    try {
      const responseSpreadsheetId = extractSpreadsheetId(setting.responseSpreadsheetUrl);
      if (!responseSpreadsheetId) {
        await writeRecruitmentStatus(setting, "回答スプレッドシートURL待ち");
        continue;
      }
      if (setting.pollChannelId && !/^\d{17,20}$/.test(setting.pollChannelId)) {
        throw new Error("投票チャンネルIDの形式が正しくありません");
      }
      if (setting.passChannelId && !/^\d{17,20}$/.test(setting.passChannelId)) {
        throw new Error("合格発表チャンネルIDの形式が正しくありません");
      }
      const responseSheetName = await resolveResponseSheetName(responseSpreadsheetId, setting.responseSheetName);
      const formResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: responseSpreadsheetId,
        range: `${quoteSheetName(responseSheetName)}!A1:Z1000`,
        valueRenderOption: "UNFORMATTED_VALUE",
      });
      const formValues = formResponse.data.values || [];
      const headers = formValues[0] || [];
      if (!headers.length) throw new Error("回答シートの見出し行が空です");

      const newRows = [];
      for (let index = 1; index < formValues.length; index += 1) {
        const responseRow = formValues[index] || [];
        if (!responseRow.some((value) => String(value || "").trim())) continue;
        const application = responseApplication(headers, responseRow);
        const sourceKey = [
          setting.roundName,
          responseSpreadsheetId,
          application.submittedAt,
          application.discordId,
          application.name,
        ].join("|");
        if (sourceKeys.has(sourceKey)) continue;
        const emptyIndex = Array.from({ length: 990 }, (_, rowIndex) => rowIndex)
          .find((rowIndex) => !occupiedRows.has(rowIndex));
        if (emptyIndex === undefined) throw new Error("応募管理シートの保存行が不足しています");
        const rowNumber = emptyIndex + 11;
        const values = applicationSheetRow(rowNumber, setting, application, sourceKey);
        newRows.push({ rowNumber, values });
        rows[emptyIndex] = values;
        occupiedRows.add(emptyIndex);
        sourceKeys.add(sourceKey);
      }

      if (newRows.length) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId,
          requestBody: {
            valueInputOption: "USER_ENTERED",
            data: newRows.map(({ rowNumber, values }) => ({
              range: `'${applicationSheetName}'!A${rowNumber}:AD${rowNumber}`,
              values: [values],
            })),
          },
        });
      }

      let pollCreatedCount = 0;
      let pollUpdatedCount = 0;
      let passAnnouncementCount = 0;
      let documentRoleCount = 0;
      const roundApplications = rows
        .map((row, index) => applicationFromSheetRow(row || [], index + 11))
        .filter((application) => application.id && application.roundName === setting.roundName);

      for (const application of roundApplications) {
        if (!application.pollMessageId) {
          if (!setting.pollChannelId) continue;
          try {
            const state = await createApplicationPoll(setting, application);
            await writeApplicationPollState(application.rowNumber, state);
            const row = rows[application.rowNumber - 11];
            row.splice(14, 10,
              state.pollStatus, state.passVotes, state.failVotes, state.totalVotes, state.verdict,
              state.pollUrl, state.pollEndsAt, state.pollMessageId, state.pollChannelId, state.processResult);
            Object.assign(application, state);
            pollCreatedCount += 1;
          } catch (error) {
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `'${applicationSheetName}'!X${application.rowNumber}`,
              valueInputOption: "RAW",
              requestBody: { values: [[`投票作成エラー: ${error.message}`]] },
            });
          }
          continue;
        }

        if (application.pollStatus !== "FINAL") {
          try {
            const state = await fetchApplicationPoll(application, setting);
            if (applicationPollStateChanged(application, state)) {
              await writeApplicationPollState(application.rowNumber, state);
              Object.assign(application, state);
              pollUpdatedCount += 1;
            }
          } catch (error) {
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `'${applicationSheetName}'!X${application.rowNumber}`,
              valueInputOption: "RAW",
              requestBody: { values: [[`投票読込エラー: ${error.message}`]] },
            });
          }
        }

        if (application.pollStatus === "FINAL") {
          if (await reconcileStep1DocumentRole(guild, setting, application, roundApplications)) documentRoleCount += 1;
        }

        if (application.pollStatus === "FINAL"
          && application.verdict === "合格"
          && setting.passChannelId
          && !application.passAnnouncementMessageId) {
          try {
            const announcement = await sendApplicationPass(setting, application);
            const result = `${application.processResult || "投票結果を確定"} / 合格発表済（書類非表示）${announcement.applicantMatched ? "" : "（本人メンション未解決）"}`;
            await writeApplicationAnnouncementState(application.rowNumber, announcement.messageId, result);
            application.passAnnouncementMessageId = announcement.messageId;
            application.processResult = result;
            // Keep the per-applicant resend button available after success.
            try {
              await setAnnouncementRetryButton(application.pollChannelId, application.pollMessageId, "document", application.id, true);
              announcementRetryButtonsInstalled.add(`doc:${application.id}`);
            } catch (buttonError) {
              console.warn(`書類合格発表の再送ボタン設置失敗 (${application.id}):`, buttonError.message);
            }
            passAnnouncementCount += 1;
          } catch (error) {
            await setAnnouncementRetryButton(application.pollChannelId, application.pollMessageId, "document", application.id, true).catch((buttonError) => {
              console.warn(`書類合格発表の再送信ボタン設置失敗 (${application.id}):`, buttonError.message);
            });
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `'${applicationSheetName}'!X${application.rowNumber}`,
              valueInputOption: "RAW",
              requestBody: { values: [[`合格発表エラー: ${error.message}`]] },
            });
          }
        } else if (application.pollStatus === "FINAL"
          && application.verdict === "合格"
          && setting.passChannelId
          && application.passAnnouncementMessageId
          && !String(application.processResult || "").includes("書類非表示")) {
          try {
            const applicantMatched = await sanitizeApplicationPass(setting, application);
            const result = `${application.processResult || "投票結果を確定"} / 書類非表示へ更新${applicantMatched ? "" : "（本人メンション未解決）"}`;
            await writeApplicationAnnouncementState(application.rowNumber, application.passAnnouncementMessageId, result);
            application.processResult = result;
            try {
              await setAnnouncementRetryButton(application.pollChannelId, application.pollMessageId, "document", application.id, true);
              announcementRetryButtonsInstalled.add(`doc:${application.id}`);
            } catch (buttonError) {
              console.warn(`書類合格発表の再送ボタン更新失敗 (${application.id}):`, buttonError.message);
            }
          } catch (error) {
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `'${applicationSheetName}'!X${application.rowNumber}`,
              valueInputOption: "RAW",
              requestBody: { values: [[`合格発表更新エラー: ${error.message}`]] },
            });
          }
        } else if (application.pollStatus === "FINAL"
          && application.verdict === "合格"
          && setting.passChannelId
          && application.passAnnouncementMessageId
          && !announcementRetryButtonsInstalled.has(`doc:${application.id}`)) {
          try {
            await setAnnouncementRetryButton(application.pollChannelId, application.pollMessageId, "document", application.id, true);
            announcementRetryButtonsInstalled.add(`doc:${application.id}`);
          } catch (error) {
            console.warn(`書類合格発表の再送ボタン更新失敗 (${application.id}):`, error.message);
          }
        }
      }

      const previewCount = roundApplications.filter((application) => application.pollStatus === "PREVIEW").length;
      const finalCount = roundApplications.filter((application) => application.pollStatus === "FINAL").length;
      const channelWait = !setting.pollChannelId ? " / 投票チャンネル待ち" : "";
      const announcementStatus = setting.passChannelId ? ` / 合格発表${passAnnouncementCount}件` : " / 合格発表なし";
      await writeRecruitmentStatus(
        setting,
        `稼働中: 応募${roundApplications.length}件 / 新規${newRows.length}件 / 投票作成${pollCreatedCount}件 / 更新${pollUpdatedCount}件 / ロール${documentRoleCount}件 / PREVIEW${previewCount}件 / FINAL${finalCount}件 / 締切${setting.pollDurationHours}時間または${setting.pollVoteLimit}票${announcementStatus}${channelWait}`,
      );
    } catch (error) {
      const message = error.response?.data?.error?.message || error.message || String(error);
      // A quota error must not trigger a second status write, which would
      // prolong the outage. The central scheduler will retry automatically.
      if (isSheetsQuotaError(error)) throw error;
      await writeRecruitmentStatus(setting, `エラー: ${message}`);
      console.error(`応募フォーム連携失敗: ${setting.roundName}`, message);
    }
  }
}

function interviewSettingFromRow(row, index, firstRow = 4, source = "legacy") {
  const requestedDuration = Number(row[6]);
  const requestedVoteLimit = Number(row[7]);
  return {
    rowNumber: index + firstRow,
    source,
    enabled: String(row[0] || "").trim() === "はい",
    roundName: String(row[1] || "").trim(),
    commandChannelId: String(row[2] || "").replace(/^'/, "").trim(),
    interviewerRoleId: String(row[3] || "").replace(/^'/, "").trim(),
    pollChannelId: String(row[4] || "").replace(/^'/, "").trim(),
    pollMessage: String(row[5] || "").trim() || "面接内容を確認し、合格または不合格へ投票してください。",
    pollDurationHours: Number.isFinite(requestedDuration) && requestedDuration >= 1 && requestedDuration <= 168
      ? Math.floor(requestedDuration)
      : 24,
    pollVoteLimit: Number.isFinite(requestedVoteLimit) && requestedVoteLimit >= 1 && requestedVoteLimit <= 1000
      ? Math.floor(requestedVoteLimit)
      : 5,
    passChannelId: String(row[8] || "").replace(/^'/, "").trim(),
    passMessage: String(row[9] || "").trim() || "面接選考の合格が決定しました。今後の案内をご確認ください。",
    questionSet: String(row[10] || "").trim() || "標準面接",
    interviewPassRoleId: String(row[15] || "").replace(/^'/, "").trim(),
  };
}

async function readInterviewSettings() {
  const unified = await unifiedRange("interview");
  if (unified) {
    return unified.rows
      .map((row, index) => interviewSettingFromRow(row, index, unified.firstRow, unified.source))
      .filter((setting) => setting.enabled);
  }
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${interviewSettingsSheetName}'!A4:P100`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return (response.data.values || [])
    .map(interviewSettingFromRow)
    .filter((setting) => setting.enabled);
}

function usableInterviewPollSetting(setting) {
  const pollChannelId = normalizedId(setting?.pollChannelId);
  return Boolean(setting?.enabled && /^\d{17,20}$/.test(pollChannelId));
}

function selectInterviewPollSetting(settings, record, interaction = null) {
  const candidates = (settings || []).filter(usableInterviewPollSetting);
  if (!candidates.length) return null;
  // Prefer the round captured on the interview record.  If an older record has
  // no matching round, fall back to the command-channel setting and finally to
  // the first valid setting so a duplicate/blank row cannot mask a real one.
  return candidates.find((setting) => setting.roundName && setting.roundName === record?.roundName)
    || candidates.find((setting) => interaction && setting.commandChannelId === interaction.channelId)
    || candidates[0];
}

async function writeInterviewStatus(setting, status) {
  const key = `${setting.source || "legacy"}:${setting.roundName}`;
  const now = Date.now();
  const previous = interviewStatusCache.get(key);
  if (previous && previous.status === status && now - previous.at < 900000) return;
  const sheetName = setting.source === "unified" ? unifiedSettingsSheetName : interviewSettingsSheetName;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetName}'!L${setting.rowNumber}:M${setting.rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[status, sheetDateTime(new Date())]] },
  });
  interviewStatusCache.set(key, { status, at: now });
}

async function readInterviewQuestions(questionSet) {
  const unified = await unifiedRange("questions");
  if (unified) {
    return unified.rows
      .map((row) => ({
        enabled: String(row[0] || "").trim() === "はい",
        questionSet: String(row[1] || "").trim(),
        id: String(row[2] || "").trim(),
        order: Number(row[3]) || 9999,
        text: String(row[4] || "").trim(),
        note: String(row[5] || "").trim(),
        required: String(row[6] || "").trim() !== "いいえ",
      }))
      .filter((question) => question.enabled && question.questionSet === questionSet && question.text)
      .sort((a, b) => a.order - b.order);
  }
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${interviewQuestionsSheetName}'!A4:G1000`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return (response.data.values || [])
    .map((row) => ({
      enabled: String(row[0] || "").trim() === "はい",
      questionSet: String(row[1] || "").trim(),
      id: String(row[2] || "").trim(),
      order: Number(row[3]) || 9999,
      text: String(row[4] || "").trim(),
      note: String(row[5] || "").trim(),
      required: String(row[6] || "").trim() !== "いいえ",
    }))
    .filter((question) => question.enabled && question.questionSet === questionSet && question.text)
    .sort((a, b) => a.order - b.order);
}

function interviewRecordFromRow(row, rowNumber) {
  return {
    rowNumber,
    id: String(row[0] || "").trim(),
    applicationId: String(row[1] || "").trim(),
    roundName: String(row[2] || "").trim(),
    applicantName: String(row[3] || "").trim(),
    applicantDiscordName: String(row[4] || "").trim(),
    applicantDiscordId: String(row[5] || "").replace(/^'/, "").trim(),
    interviewerId: String(row[6] || "").replace(/^'/, "").trim(),
    interviewerName: String(row[7] || "").trim(),
    startedAt: row[8] || "",
    completedAt: row[9] || "",
    questionSet: String(row[10] || "").trim(),
    questionSnapshot: String(row[11] || "").trim(),
    answerMemo: String(row[12] || "").trim(),
    interviewStatus: String(row[13] || "").trim(),
    pollStatus: String(row[14] || "").trim(),
    passVotes: Number(row[15]) || 0,
    failVotes: Number(row[16]) || 0,
    totalVotes: Number(row[17]) || 0,
    verdict: String(row[18] || "").trim(),
    pollUrl: String(row[19] || "").trim(),
    pollEndsAt: row[20] || "",
    pollMessageId: String(row[21] || "").replace(/^'/, "").trim(),
    pollChannelId: String(row[22] || "").replace(/^'/, "").trim(),
    processResult: String(row[23] || "").trim(),
    passAnnouncementMessageId: String(row[24] || "").replace(/^'/, "").trim(),
    updatedAt: row[25] || "",
    calledAt: row[26] || "",
    callMessageId: String(row[27] || "").replace(/^'/, "").trim(),
    calledBy: String(row[28] || "").trim(),
    roleStatus: String(row[29] || "").trim(),
  };
}

const shortReadCache = new Map();
const interviewUiSessions = new Map();
const interviewUiSessionTtlMs = 15 * 60 * 1000;

function cacheInterviewSession(record) {
  if (interviewUiSessions.size >= 500) {
    const now = Date.now();
    for (const [id, cached] of interviewUiSessions) {
      if (cached.expiresAt <= now) interviewUiSessions.delete(id);
    }
  }
  if (record?.id) interviewUiSessions.set(record.id, { record, expiresAt: Date.now() + interviewUiSessionTtlMs });
  return record;
}

function cachedInterviewSession(id) {
  const cached = interviewUiSessions.get(id);
  if (!cached) return null;
  if (cached.expiresAt <= Date.now()) {
    interviewUiSessions.delete(id);
    return null;
  }
  cached.expiresAt = Date.now() + interviewUiSessionTtlMs;
  return cached.record;
}
function invalidateShortRead(key) {
  shortReadCache.delete(key);
}
async function cachedShortRead(key, ttlMs, loader) {
  const now = Date.now();
  const cached = shortReadCache.get(key);
  if (cached?.value && cached.expiresAt > now) return cached.value;
  if (cached?.promise) return cached.promise;
  const promise = loader()
    .then((value) => {
      shortReadCache.set(key, { value, expiresAt: Date.now() + ttlMs });
      return value;
    })
    .catch((error) => {
      shortReadCache.delete(key);
      throw error;
    });
  shortReadCache.set(key, { promise, value: cached?.value, expiresAt: cached?.expiresAt || 0 });
  return promise;
}

async function readInterviewRecords() {
  return cachedShortRead("interview-records", 3000, async () => {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${interviewManagementSheetName}'!A11:AD1000`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    const records = (response.data.values || [])
      .map((row, index) => interviewRecordFromRow(row, index + 11))
      .filter((record) => record.id);
    records.forEach(cacheInterviewSession);
    return records;
  });
}

async function readStep1Applications() {
  return cachedShortRead("step1-applications", 3000, async () => {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${applicationSheetName}'!A11:AD1000`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    return (response.data.values || [])
      .map((row, index) => applicationFromSheetRow(row, index + 11))
      .filter((application) => application.id);
  });
}

async function eligibleInterviewApplications(setting, search = "") {
  const [applications, records] = await Promise.all([readStep1Applications(), readInterviewRecords()]);
  const activeRecords = records.filter((record) => record.interviewStatus !== "CANCELLED");
  const reserved = new Set(activeRecords
    .map((record) => record.applicationId));
  const reservedApplicants = new Set(activeRecords.map((record) => (
    record.applicantDiscordId || normalizeDiscordUsername(record.applicantDiscordName)
  )).filter(Boolean));
  const needle = String(search || "").trim().toLowerCase();
  const candidates = applications.filter((application) => (
    application.roundName === setting.roundName
    && application.pollStatus === "FINAL"
    && application.verdict === "合格"
    && application.documentRoleStatus.startsWith("書類合格ロール付与済")
    && (!needle || [application.id, application.name, application.discordId]
      .some((value) => String(value || "").toLowerCase().includes(needle)))
  ));
  const latestByApplicant = new Map();
  for (const application of candidates) {
    const key = application.resolvedDiscordId || normalizeDiscordUsername(application.discordId) || application.id;
    latestByApplicant.set(key, application);
  }
  return [...latestByApplicant.entries()]
    .filter(([key, application]) => !reserved.has(application.id) && !reservedApplicants.has(key))
    .map(([, application]) => application);
}

function interviewQuestionSnapshot(questions) {
  return JSON.stringify(questions.map((question, index) => ({
    id: question.id || `Q${String(index + 1).padStart(3, "0")}`,
    order: index + 1,
    text: question.text,
    note: question.note || "",
    required: Boolean(question.required),
  })));
}

function interviewSnapshotQuestions(record) {
  try {
    const parsed = JSON.parse(record.questionSnapshot || "[]");
    if (Array.isArray(parsed) && parsed.length) return parsed;
  } catch {
    // 旧形式の面接レコードは、番号付きの質問一覧として読み替える。
  }
  return String(record.questionSnapshot || "")
    .split(/\n(?=\d+\.\s)/)
    .map((text, index) => ({ id: `LEGACY-${index + 1}`, order: index + 1, text: text.replace(/^\d+\.\s*(?:【[^】]+】\s*)?/, "").trim(), note: "", required: true }))
    .filter((question) => question.text);
}

function interviewAnswers(record) {
  try {
    const parsed = JSON.parse(record.answerMemo || "{}");
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) return parsed;
  } catch {
    // 旧形式の自由記述メモは先頭質問の回答として保持する。
  }
  const questions = interviewSnapshotQuestions(record);
  return record.answerMemo && questions[0] ? { [questions[0].id]: record.answerMemo } : {};
}

function interviewReady(record) {
  const answers = interviewAnswers(record);
  const required = interviewSnapshotQuestions(record).filter((question) => question.required);
  return required.length > 0 && required.every((question) => String(answers[question.id] || "").trim());
}

function interviewAnswerSummary(record) {
  const questions = interviewSnapshotQuestions(record);
  const answers = interviewAnswers(record);
  return questions.map((question, index) => (
    `**Q${index + 1}. ${truncateDiscord(question.text, 240)}**\n${truncateDiscord(answers[question.id] || "未回答", 500)}`
  )).join("\n\n");
}

function nextInterviewId(applicationId, records) {
  const count = records.filter((record) => record.applicationId === applicationId).length + 1;
  return `INT-${applicationId}-${String(count).padStart(2, "0")}`;
}

async function reserveInterview(setting, application, interviewer, questions) {
  const records = await readInterviewRecords();
  if (records.some((record) => record.applicationId === application.id && record.interviewStatus !== "CANCELLED")) {
    throw new Error("この応募者はすでに別の面接官が選択しています。");
  }
  const occupiedRows = new Set(records.map((record) => record.rowNumber));
  const emptyIndex = Array.from({ length: 990 }, (_, index) => index)
    .find((index) => !occupiedRows.has(index + 11));
  if (emptyIndex === undefined) throw new Error("面接管理シートの保存行が不足しています。");
  const rowNumber = emptyIndex + 11;
  const id = nextInterviewId(application.id, records);
  const applicantDiscordId = application.resolvedDiscordId || await resolveApplicantDiscordUserId(application.discordId);
  const snapshot = interviewQuestionSnapshot(questions);
  const now = sheetDateTime(new Date());
  const values = [
    id,
    application.id,
    application.roundName,
    application.name,
    application.discordId,
    applicantDiscordId ? discordIdCell(applicantDiscordId) : "",
    discordIdCell(interviewer.id),
    interviewer.displayName || interviewer.user?.username || interviewer.id,
    now,
    "",
    setting.questionSet,
    snapshot,
    "",
    "IN_PROGRESS",
    "",
    0,
    0,
    0,
    "",
    "",
    "",
    "",
    "",
    "質問回答待ち",
    "",
    now,
    "",
    "",
    "",
    "",
  ];
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${interviewManagementSheetName}'!A${rowNumber}:AD${rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [values] },
  });
  invalidateShortRead("interview-records");
  return cacheInterviewSession(interviewRecordFromRow(values, rowNumber));
}

function canInterview(interaction, setting) {
  if (interaction.guildId !== guildId || interaction.channelId !== setting.commandChannelId) return false;
  if (setting.interviewerRoleId) return Boolean(interaction.member?.roles?.cache?.has(setting.interviewerRoleId));
  // ロール未設定時は、設定済みの非公開チャンネルにアクセスできる人を面接官として扱う。
  // 初回導入時にロールIDの登録を必須にせず、必要になった時点で制限を追加できる。
  return true;
}

function interviewSessionComponents(record, questionIndex = 0) {
  const questions = interviewSnapshotQuestions(record);
  const safeIndex = Math.max(0, Math.min(questionIndex, Math.max(questions.length - 1, 0)));
  const answers = interviewAnswers(record);
  const current = questions[safeIndex];
  const jump = new StringSelectMenuBuilder()
    .setCustomId(`iv:jump:${record.id}`)
    .setPlaceholder(`質問を選択（${safeIndex + 1}/${questions.length}）`)
    .addOptions(questions.slice(0, 24).map((question, index) => new StringSelectMenuOptionBuilder()
      .setLabel(truncateDiscord(`${answers[question.id] ? "✅" : question.required ? "必須" : "任意"} Q${index + 1} ${question.text}`, 100))
      .setValue(String(index))
      .setDefault(index === safeIndex)));
  return [
    new ActionRowBuilder().addComponents(jump),
    new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`iv:page:${record.id}:${Math.max(safeIndex - 1, 0)}`)
      .setLabel("← 前へ")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safeIndex === 0),
    new ButtonBuilder()
      .setCustomId(`iv:q:${record.id}:${safeIndex}`)
      .setLabel(current && answers[current.id] ? "回答を修正" : "回答する")
      .setStyle(ButtonStyle.Primary),
    new ButtonBuilder()
      .setCustomId(`iv:page:${record.id}:${Math.min(safeIndex + 1, Math.max(questions.length - 1, 0))}`)
      .setLabel("次へ →")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safeIndex >= questions.length - 1),
    ),
    new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`iv:vote:${record.id}`)
      .setLabel("回答完了・投票へ")
      .setStyle(ButtonStyle.Success)
      .setDisabled(!interviewReady(record)),
    new ButtonBuilder()
      .setCustomId(`iv:cancel:${record.id}`)
      .setLabel("面接を中止")
      .setStyle(ButtonStyle.Danger),
    ),
  ];
}

function interviewSessionEmbed(record, questionIndex = 0) {
  const questions = interviewSnapshotQuestions(record);
  const answers = interviewAnswers(record);
  const safeIndex = Math.max(0, Math.min(questionIndex, Math.max(questions.length - 1, 0)));
  const answered = questions.filter((question) => String(answers[question.id] || "").trim()).length;
  const required = questions.filter((question) => question.required);
  const requiredAnswered = required.filter((question) => String(answers[question.id] || "").trim()).length;
  const current = questions[safeIndex];
  const questionList = questions.slice(0, 24).map((question, index) =>
    `${index === safeIndex ? "▶" : answers[question.id] ? "✅" : question.required ? "🔸" : "▫️"} Q${index + 1} ${truncateDiscord(question.text, 34)}`
  ).join("\n");
  const currentValue = current
    ? `${current.text}${current.note ? `\n\n補足: ${current.note}` : ""}${answers[current.id] ? `\n\n**保存済み回答**\n${answers[current.id]}` : "\n\n*未回答*"}`
    : "質問がありません。";
  return new EmbedBuilder()
    .setColor(interviewReady(record) ? 0x16a34a : 0x2563eb)
    .setTitle(truncateDiscord(`STEP2 面接進行｜${record.applicantName || record.applicationId}`, 256))
    .setDescription(`応募ID: **${record.applicationId}**　面接ID: **${record.id}**\n進捗: **必須 ${requiredAnswered}/${required.length}｜全体 ${answered}/${questions.length}**`)
    .addFields(
      { name: `現在の質問｜Q${safeIndex + 1}${current?.required ? "【必須】" : "【任意】"}`, value: truncateDiscord(currentValue, 1024) },
      { name: "質問一覧", value: truncateDiscord(questionList, 1024) || "質問なし" },
    )
    .setFooter({ text: interviewReady(record) ? "必須回答が完了しました。投票へ進めます。" : "質問を選び、1問ずつ回答してください。回答は自動保存されます。" });
}

async function interviewRecordById(id) {
  return cachedInterviewSession(id) || (await readInterviewRecords()).find((record) => record.id === id) || null;
}

async function updateInterviewRecord(record, valuesByColumn) {
  const data = Object.entries(valuesByColumn).map(([column, value]) => ({
    range: `'${interviewManagementSheetName}'!${column}${record.rowNumber}`,
    values: [[value]],
  }));
  data.push({ range: `'${interviewManagementSheetName}'!Z${record.rowNumber}`, values: [[sheetDateTime(new Date())]] });
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: "USER_ENTERED", data },
  });
  invalidateShortRead("interview-records");
  Object.entries(valuesByColumn).forEach(([column, value]) => {
    const fields = { M: "answerMemo", N: "interviewStatus", O: "pollStatus", X: "processResult" };
    if (fields[column]) record[fields[column]] = value;
  });
  cacheInterviewSession(record);
}

async function settingForInteraction(interaction) {
  const settings = await readInterviewSettings();
  return settings.find((setting) => setting.commandChannelId === interaction.channelId) || null;
}

async function handleInterviewCommand(interaction) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  const setting = await settingForInteraction(interaction);
  if (!setting) {
    await interaction.editReply("このチャンネルは「面接設定」で有効化されていません。");
    return;
  }
  if (!canInterview(interaction, setting)) {
    await interaction.editReply("このコマンドを実行できる面接官ロールがありません。");
    return;
  }
  const activeRecord = (await readInterviewRecords())
    .filter((record) => record.interviewerId === interaction.user.id
      && ["IN_PROGRESS", "READY"].includes(record.interviewStatus))
    .sort((left, right) => right.rowNumber - left.rowNumber)[0];
  if (activeRecord) {
    await interaction.editReply({
      content: "進行中の面接を再開しました。",
      embeds: [interviewSessionEmbed(activeRecord, 0)],
      components: interviewSessionComponents(activeRecord, 0),
    });
    return;
  }
  const search = interaction.options?.getString?.("search") || "";
  const candidates = await eligibleInterviewApplications(setting, search);
  if (!candidates.length) {
    await interaction.editReply(search
      ? `「${search}」に一致する面接待ちの書類合格者はいません。`
      : "現在、面接待ちの書類合格者はいません。");
    return;
  }
  const shown = candidates.slice(0, 25);
  const select = new StringSelectMenuBuilder()
    .setCustomId(`iv:select:${setting.rowNumber}`)
    .setPlaceholder("面接する応募者を選択")
    .addOptions(shown.map((application) => new StringSelectMenuOptionBuilder()
      .setLabel(truncateDiscord(`${application.name || application.discordId} — ${application.id}`, 100))
      .setDescription(truncateDiscord(`${application.discordId} / ${application.roundName}`, 100))
      .setValue(application.id)));
  const description = candidates.length > 25
    ? `候補者${candidates.length}人のうち先頭25人を表示しています。/mensetu の search で名前・応募IDを絞り込めます。`
    : `書類合格通知済みの候補者${candidates.length}人から選択してください。`;
  await interaction.editReply({
    embeds: [new EmbedBuilder().setColor(0x2563eb).setTitle("STEP2 面接対象者を選択").setDescription(description)],
    components: [new ActionRowBuilder().addComponents(select)],
  });
}

async function handleInterviewSelect(interaction) {
  const setting = await settingForInteraction(interaction);
  if (!setting || !canInterview(interaction, setting)) throw new Error("面接設定または実行権限を確認できません。");
  const applicationId = interaction.values[0];
  const candidates = await eligibleInterviewApplications(setting, applicationId);
  const application = candidates.find((candidate) => candidate.id === applicationId);
  if (!application) throw new Error("この応募者は現在面接対象ではありません。再度 /mensetu を実行してください。");
  const questions = await readInterviewQuestions(setting.questionSet);
  if (!questions.length) throw new Error(`質問セット「${setting.questionSet}」に有効な質問がありません。`);
  if (questions.length > 24) throw new Error("1つの質問セットに登録できる有効な質問は24問までです。");
  const record = await reserveInterview(setting, application, interaction.member, questions);
  await interaction.editReply({
    embeds: [interviewSessionEmbed(record, 0)],
    components: interviewSessionComponents(record, 0),
  });
}

async function handleInterviewPageButton(interaction, interviewId, questionIndex) {
  const record = await interviewRecordById(interviewId);
  if (!record || record.interviewerId !== interaction.user.id) throw new Error("この面接を操作できません。");
  if (!["IN_PROGRESS", "READY"].includes(record.interviewStatus)) throw new Error("この面接はすでに投票段階へ進んでいます。");
  await interaction.editReply({ embeds: [interviewSessionEmbed(record, questionIndex)], components: interviewSessionComponents(record, questionIndex) });
}

async function handleInterviewJump(interaction, interviewId, questionIndex) {
  return handleInterviewPageButton(interaction, interviewId, questionIndex);
}

async function handleInterviewQuestionButton(interaction, interviewId, questionIndex) {
  const record = await interviewRecordById(interviewId);
  if (!record || record.interviewerId !== interaction.user.id) {
    await interaction.reply({ content: "この面接を操作できる面接官ではありません。", flags: MessageFlags.Ephemeral });
    return;
  }
  if (!["IN_PROGRESS", "READY"].includes(record.interviewStatus)) {
    await interaction.reply({ content: "この面接はすでに投票段階へ進んでいます。", flags: MessageFlags.Ephemeral });
    return;
  }
  const questions = interviewSnapshotQuestions(record);
  const safeIndex = Math.max(0, Math.min(questionIndex, questions.length - 1));
  const question = questions[safeIndex];
  if (!question) throw new Error("質問が見つかりません。");
  const answers = interviewAnswers(record);
  const input = new TextInputBuilder()
    .setCustomId("answer")
    .setLabel(`Q${safeIndex + 1} の回答${question.required ? "（必須）" : "（任意）"}`)
    .setStyle(TextInputStyle.Paragraph)
    .setRequired(question.required)
    .setMaxLength(1500)
    .setPlaceholder(truncateDiscord(question.text, 100));
  if (answers[question.id]) input.setValue(String(answers[question.id]).slice(0, 1500));
  await interaction.showModal(new ModalBuilder()
    .setCustomId(`iv:answer:${record.id}:${safeIndex}`)
    .setTitle(`面接質問 Q${safeIndex + 1}/${questions.length}`)
    .addComponents(new ActionRowBuilder().addComponents(input)));
}

async function handleInterviewAnswerModal(interaction, interviewId, questionIndex) {
  const record = await interviewRecordById(interviewId);
  if (!record || record.interviewerId !== interaction.user.id) throw new Error("この面接を操作できません。");
  if (!["IN_PROGRESS", "READY"].includes(record.interviewStatus)) throw new Error("この面接は回答を変更できません。");
  const questions = interviewSnapshotQuestions(record);
  const safeIndex = Math.max(0, Math.min(questionIndex, questions.length - 1));
  const question = questions[safeIndex];
  if (!question) throw new Error("質問が見つかりません。");
  const answer = interaction.fields.getTextInputValue("answer").trim();
  if (question.required && !answer) throw new Error("必須質問には回答を入力してください。");
  const answers = interviewAnswers(record);
  if (answer) answers[question.id] = answer;
  else delete answers[question.id];
  const serialized = JSON.stringify(answers);
  record.answerMemo = serialized;
  const ready = interviewReady(record);
  const nextUnanswered = questions.findIndex((item, index) => index > safeIndex && item.required && !answers[item.id]);
  const nextIndex = nextUnanswered >= 0 ? nextUnanswered : Math.min(safeIndex + 1, questions.length - 1);
  const status = ready ? "READY" : "IN_PROGRESS";
  await updateInterviewRecord(record, { M: serialized, N: status, X: ready ? "必須質問の回答完了" : "質問回答を自動保存" });
  Object.assign(record, { interviewStatus: status, processResult: ready ? "必須質問の回答完了" : "質問回答を自動保存" });
  await interaction.editReply({ embeds: [interviewSessionEmbed(record, nextIndex)], components: interviewSessionComponents(record, nextIndex) });
}

function interviewPollEmbed(record) {
  return new EmbedBuilder()
    .setColor(0x7c3aed)
    .setTitle(truncateDiscord(`STEP2 面接審査｜${record.applicantName || record.applicationId}`, 256))
    .setDescription(`応募ID: **${record.applicationId}**\n面接官: **${record.interviewerName}**`)
    .addFields(
      { name: "面接質問と回答", value: truncateDiscord(interviewAnswerSummary(record), 1024) },
    )
    .setTimestamp();
}

async function createInterviewPoll(setting, record) {
  const channel = await textChannel(setting.pollChannelId, "面接投票");
  const message = await channel.send({
    content: setting.pollMessage,
    embeds: [interviewPollEmbed(record)],
    poll: {
      question: { text: `${truncateDiscord(record.applicantName || record.applicationId, 220)} を面接合格としますか？` },
      answers: [
        { text: "合格", emoji: "✅" },
        { text: "不合格", emoji: "❌" },
      ],
      allowMultiselect: false,
      duration: setting.pollDurationHours,
    },
    allowedMentions: { parse: [] },
  });
  return message;
}

async function handleInterviewVoteButton(interaction, interviewId) {
  const record = await interviewRecordById(interviewId);
  if (!record || record.interviewerId !== interaction.user.id) throw new Error("この面接を操作できません。");
  if (record.interviewStatus !== "READY" || !interviewReady(record)) throw new Error("必須質問へすべて回答してください。");
  const settings = await readInterviewSettings();
  const setting = selectInterviewPollSetting(settings, record, interaction);
  if (!setting?.pollChannelId) throw new Error("面接設定の投票チャンネルIDが未設定です。");
  const message = await createInterviewPoll(setting, record);
  const pollEndsAt = message.poll?.expiresTimestamp ? sheetDateTime(new Date(message.poll.expiresTimestamp)) : "";
  await updateInterviewRecord(record, {
    J: sheetDateTime(new Date()),
    N: "PREVIEW",
    O: "PREVIEW",
    P: 0,
    Q: 0,
    R: 0,
    S: "投票待ち",
    T: message.url,
    U: pollEndsAt,
    V: discordIdCell(message.id),
    W: discordIdCell(message.channelId),
    X: "面接投票を開始",
  });
  await interaction.editReply({
    content: `面接回答を保存し、面接投票を開始しました。\n${message.url}`,
    embeds: [],
    components: [],
  });
}

async function handleInterviewCancelButton(interaction, interviewId) {
  const record = await interviewRecordById(interviewId);
  if (!record || record.interviewerId !== interaction.user.id) throw new Error("この面接を操作できません。");
  if (!["IN_PROGRESS", "READY"].includes(record.interviewStatus)) throw new Error("投票開始後の面接はキャンセルできません。");
  await updateInterviewRecord(record, { N: "CANCELLED", X: "面接官がキャンセル" });
  await interaction.editReply({ content: "面接をキャンセルしました。応募者は再び候補一覧へ戻ります。", embeds: [], components: [] });
}

async function sendInterviewPass(setting, record, { resend = false } = {}) {
  if (!setting.passChannelId) throw new Error("面接合格発表チャンネルIDが未設定です");
  const channel = await textChannel(setting.passChannelId, "面接合格発表");
  const userId = record.applicantDiscordId || await resolveApplicantDiscordUserId(record.applicantDiscordName);
  const content = [setting.passMessage, userId ? `<@${userId}>` : ""].filter(Boolean).join("\n");
  const message = await channel.send({
    content,
    allowedMentions: { users: userId ? [userId] : [] },
    nonce: discordNonce("ivpass", resend ? `${record.id}:${Date.now()}` : record.id),
    enforceNonce: true,
  });
  return { messageId: message.id, applicantMatched: Boolean(userId) };
}

async function canResendAnnouncement(interaction) {
  if (interaction.memberPermissions?.has(PermissionFlagsBits.Administrator)
    || interaction.memberPermissions?.has(PermissionFlagsBits.ManageMessages)) return true;
  const settings = await readRankOperationSettings();
  return Boolean(settings.managementRoleId && interaction.member?.roles?.cache?.has(settings.managementRoleId));
}

async function handleDocumentAnnouncementResend(interaction, applicationId) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  if (!await canResendAnnouncement(interaction)) {
    await interaction.editReply("合格発表の再送信には管理権限が必要です。");
    return;
  }
  const applications = await readStep1Applications();
  const application = applications.find((item) => item.id === applicationId);
  if (application && (interaction.channelId !== application.pollChannelId || interaction.message?.id !== application.pollMessageId)) {
    await interaction.editReply("この再送信ボタンは現在の投票メッセージと一致しません。最新の投票画面を確認してください。");
    return;
  }
  if (!application || application.pollStatus !== "FINAL" || application.verdict !== "合格") {
    await interaction.editReply("この応募は現在、再送信できる合格状態ではありません。シートの投票結果を確認してください。");
    return;
  }
  const settings = await readRecruitmentSettings();
  const setting = settings.find((item) => item.roundName === application.roundName);
  if (!setting?.passChannelId) {
    await interaction.editReply("合格発表チャンネルが未設定です。設定シートを確認してください。");
    return;
  }
  try {
    const announcement = await sendApplicationPass(setting, application, { resend: true });
    const result = `${application.processResult || "投票結果を確定"} / 合格発表再送済（書類非表示）${announcement.applicantMatched ? "" : "（本人メンション未解決）"}`;
    await writeApplicationAnnouncementState(application.rowNumber, announcement.messageId, result);
    await setAnnouncementRetryButton(application.pollChannelId, application.pollMessageId, "document", application.id, true).catch(() => {});
    await interaction.editReply(`書類合格発表を再送信しました。${announcement.applicantMatched ? "本人メンション付きです。" : "本人メンションは解決できませんでした。"}`);
  } catch (error) {
    await setAnnouncementRetryButton(application.pollChannelId, application.pollMessageId, "document", application.id, true).catch(() => {});
    await interaction.editReply(`再送信に失敗しました: ${error.message}\n設定とBot権限を確認して、もう一度押してください。`);
  }
}

async function handleInterviewAnnouncementResend(interaction, interviewId) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  if (!await canResendAnnouncement(interaction)) {
    await interaction.editReply("合格発表の再送信には管理権限が必要です。");
    return;
  }
  const [records, settings] = await Promise.all([readInterviewRecords(), readInterviewSettings()]);
  const record = records.find((item) => item.id === interviewId);
  if (record && (interaction.channelId !== record.pollChannelId || interaction.message?.id !== record.pollMessageId)) {
    await interaction.editReply("この再送信ボタンは現在の投票メッセージと一致しません。最新の投票画面を確認してください。");
    return;
  }
  if (!record || record.pollStatus !== "FINAL" || record.verdict !== "合格") {
    await interaction.editReply("この面接は現在、再送信できる合格状態ではありません。投票結果を確認してください。");
    return;
  }
  const setting = settings.find((item) => item.roundName === record.roundName);
  if (!setting?.passChannelId) {
    await interaction.editReply("面接合格発表チャンネルが未設定です。設定シートを確認してください。");
    return;
  }
  try {
    const announcement = await sendInterviewPass(setting, record, { resend: true });
    const result = `${record.processResult || "面接結果を確定"} / 面接合格発表再送済${announcement.applicantMatched ? "" : "（本人メンション未解決）"}`;
    await updateInterviewRecord(record, { X: result, Y: discordIdCell(announcement.messageId) });
    await setAnnouncementRetryButton(record.pollChannelId, record.pollMessageId, "interview", record.id, true).catch(() => {});
    await interaction.editReply(`面接合格発表を再送信しました。${announcement.applicantMatched ? "本人メンション付きです。" : "本人メンションは解決できませんでした。"}`);
  } catch (error) {
    await setAnnouncementRetryButton(record.pollChannelId, record.pollMessageId, "interview", record.id, true).catch(() => {});
    await interaction.editReply(`再送信に失敗しました: ${error.message}\n設定とBot権限を確認して、もう一度押してください。`);
  }
}

async function readOnboardingRows() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${onboardingSheetName}'!A2:P500`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const values = response.data.values || [];
  const headers = values[0] || onboardingHeaders;
  return {
    headers,
    headerMap: new Map(headers.map((header, index) => [String(header), index])),
    rows: values.slice(1),
  };
}

function onboardingCellRange(onboardingSheet, header, rowNumber) {
  const index = onboardingSheet.headerMap.get(header);
  if (index === undefined) throw new Error(`${onboardingSheetName}シートに「${header}」列がありません。`);
  return `'${onboardingSheetName}'!${columnLetter(index)}${rowNumber}`;
}

function onboardingRowUpdate(onboardingSheet, rowNumber, fields) {
  return Object.entries(fields).map(([header, value]) => ({
    range: onboardingCellRange(onboardingSheet, header, rowNumber),
    values: [[value]],
  }));
}

async function upsertOnboardingRecord(record) {
  const onboardingSheet = await readOnboardingRows();
  const applicationColumn = onboardingSheet.headerMap.get("応募ID");
  const interviewColumn = onboardingSheet.headerMap.get("面接ID");
  const existingIndex = onboardingSheet.rows.findIndex((row) => (
    String(row[applicationColumn] || "").trim() === record.applicationId
    || String(row[interviewColumn] || "").trim() === record.id
  ));
  if (existingIndex >= 0) return existingIndex + 3;
  const emptyIndex = onboardingSheet.rows.findIndex((row) => !String(row[applicationColumn] || "").trim());
  const rowNumber = emptyIndex >= 0 ? emptyIndex + 3 : onboardingSheet.rows.length + 3;
  const fields = {
    "応募ID": record.applicationId,
    "面接ID": record.id,
    "募集回": record.roundName,
    "受験者名": record.applicantName,
    "Discordユーザー名": record.applicantDiscordName,
    "DiscordユーザーID": record.applicantDiscordId ? discordIdCell(record.applicantDiscordId) : "",
    "面接合格日時": sheetDateTime(new Date()),
    "手続き完了": false,
    "採用状態": "備品手続き待ち",
    "ロール処理結果": record.roleStatus || "面接合格ロール処理待ち",
    "呼出日時": record.calledAt || "",
    "呼出者": record.calledBy || "",
  };
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: onboardingRowUpdate(onboardingSheet, rowNumber, fields),
    },
  });
  return rowNumber;
}

async function reconcileInterviewDecision(guild, setting, recruitmentSetting, record) {
  if (record.pollStatus !== "FINAL" || !["合格", "不合格"].includes(record.verdict)) return false;
  if (record.roleStatus.startsWith("採用完了")) return false;
  const documentRoleId = configuredRoleId(guild, recruitmentSetting?.documentPassRoleId, ["書類合格者", "書類合格"]);
  const interviewRoleId = configuredRoleId(guild, setting.interviewPassRoleId, ["面接合格者", "面接合格"]);
  if (record.verdict === "合格" && !interviewRoleId) {
    const status = "面接合格ロール設定待ち";
    if (record.roleStatus !== status) await updateInterviewRecord(record, { AD: status });
    return false;
  }
  try {
    const userId = record.applicantDiscordId || await resolveApplicantDiscordUserId(record.applicantDiscordName);
    if (!userId) throw new Error("応募者のDiscordユーザーを特定できません");
    const member = guild.members.cache.get(userId) || await guild.members.fetch(userId);
    const result = record.verdict === "合格"
      ? await reconcileMemberRoles(guild, member, { add: [interviewRoleId], remove: [documentRoleId] }, `STEP2面接合格: ${record.id}`)
      : await reconcileMemberRoles(guild, member, { remove: [documentRoleId, interviewRoleId] }, `STEP2面接不合格: ${record.id}`);
    const status = record.verdict === "合格"
      ? `面接結果反映済（面接合格付与・書類合格解除）${result.added.length || result.removed.length ? `: +${result.added.join(",")} / -${result.removed.join(",")}` : ""}`
      : `面接結果反映済（不合格・選考ロール解除）${result.removed.length ? `: -${result.removed.join(",")}` : ""}`;
    if (record.roleStatus !== status || record.applicantDiscordId !== userId) {
      await updateInterviewRecord(record, { F: discordIdCell(userId), AD: status });
      record.applicantDiscordId = userId;
      record.roleStatus = status;
    }
    const applications = await readStep1Applications();
    const linkedApplications = applications.filter((item) => (
      item.roundName === record.roundName
      && (item.id === record.applicationId
        || item.resolvedDiscordId === userId
        || normalizeDiscordUsername(item.discordId) === normalizeDiscordUsername(record.applicantDiscordName))
    ));
    for (const application of linkedApplications) {
      if (application.documentRoleStatus !== status || application.resolvedDiscordId !== userId) {
        await updateApplicationRoleState(application, userId, status);
      }
    }
    if (record.verdict === "合格") await upsertOnboardingRecord(record);
    return true;
  } catch (error) {
    const status = `面接ロールエラー: ${error.message}`;
    if (record.roleStatus !== status) await updateInterviewRecord(record, { AD: status });
    record.roleStatus = status;
    return false;
  }
}

async function processOnboarding() {
  let onboardingSheet;
  try {
    onboardingSheet = await readOnboardingRows();
  } catch (error) {
    if (error.code === 400 || /Unable to parse range|not found/i.test(error.message || "")) return;
    throw error;
  }
  const completeColumn = onboardingSheet.headerMap.get("手続き完了");
  const handlerColumn = onboardingSheet.headerMap.get("対応署員");
  const completedAtColumn = onboardingSheet.headerMap.get("完了日時");
  const discordIdColumn = onboardingSheet.headerMap.get("DiscordユーザーID");
  const roundColumn = onboardingSheet.headerMap.get("募集回");
  const applicationIdColumn = onboardingSheet.headerMap.get("応募ID");
  const interviewIdColumn = onboardingSheet.headerMap.get("面接ID");
  const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);
  const [rankMap, interviewSettings] = await Promise.all([readRankMap(), readInterviewSettings()]);
  const trialRank = rankByName(rankMap, "体験");

  for (let index = 0; index < onboardingSheet.rows.length; index += 1) {
    const row = onboardingSheet.rows[index];
    if (!isChecked(row[completeColumn]) || row[completedAtColumn]) continue;
    const rowNumber = index + 3;
    const handler = String(row[handlerColumn] || "").trim();
    if (!handler) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: { valueInputOption: "USER_ENTERED", data: onboardingRowUpdate(onboardingSheet, rowNumber, {
          "採用状態": "担当者待ち",
          "最終エラー": "対応署員を入力してください",
        }) },
      });
      continue;
    }
    try {
      if (!trialRank) throw new Error("ランク設定に有効な体験ロールがありません");
      const discordId = String(row[discordIdColumn] || "").replace(/^'/, "").trim();
      if (!discordId) throw new Error("DiscordユーザーIDが空です");
      const setting = interviewSettings.find((item) => !item.roundName || item.roundName === String(row[roundColumn] || "").trim());
      const interviewRoleId = configuredRoleId(guild, setting?.interviewPassRoleId, ["面接合格者", "面接合格"]);
      if (!interviewRoleId) throw new Error("面接合格ロールが見つかりません");
      const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId);
      const policeRoleId = [...rosterRoleIds][0];
      const result = await reconcileMemberRoles(guild, member, {
        add: [trialRank.roleId, policeRoleId],
        remove: [interviewRoleId],
      }, `採用手続き完了: ${row[0]}`);
      await syncMember(await guild.members.fetch(discordId));
      const [interviewRecords, applications] = await Promise.all([readInterviewRecords(), readStep1Applications()]);
      const interviewId = String(row[interviewIdColumn] || "").trim();
      const applicationId = String(row[applicationIdColumn] || "").trim();
      const interviewRecord = interviewRecords.find((item) => item.id === interviewId);
      const application = applications.find((item) => item.id === applicationId);
      const completionStatus = "採用完了（体験・Police Officer付与済）";
      if (interviewRecord) await updateInterviewRecord(interviewRecord, { AD: completionStatus });
      if (application) await updateApplicationRoleState(application, discordId, completionStatus);
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: { valueInputOption: "USER_ENTERED", data: onboardingRowUpdate(onboardingSheet, rowNumber, {
          "完了日時": sheetDateTime(new Date()),
          "採用状態": "採用完了",
          "ロール処理結果": `完了: +${result.added.join(", ") || "設定済み"} / -${result.removed.join(", ") || "解除済み"}`,
          "最終エラー": "",
        }) },
      });
      console.log(`採用手続き完了: ${member.displayName}`);
    } catch (error) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: { valueInputOption: "USER_ENTERED", data: onboardingRowUpdate(onboardingSheet, rowNumber, {
          "採用状態": "要確認",
          "最終エラー": error.message,
          "ロール処理結果": `エラー: ${error.message}`,
        }) },
      });
    }
  }
}

async function processInterviewPolls() {
  const now = Date.now();
  if (now - lastInterviewPollAt < interviewPollIntervalMs) return;
  lastInterviewPollAt = now;
  let settings;
  try {
    settings = await readInterviewSettings();
  } catch (error) {
    if (error.code === 400 || /Unable to parse range|not found/i.test(error.message || "")) return;
    throw error;
  }
  if (!settings.length) return;
  const [records, recruitmentSettings, guild] = await Promise.all([
    readInterviewRecords(),
    readRecruitmentSettings(),
    client.guilds.cache.get(guildId) || client.guilds.fetch(guildId),
  ]);
  for (const setting of settings) {
    // 面接設定と同じ募集回だけを処理する。別回の設定を流用すると、
    // 投票先・発表先・締切条件が別回へ誤適用されるため許可しない。
    const targetRecords = records.filter((record) => record.roundName === setting.roundName);
    let updatedCount = 0;
    let announcedCount = 0;
    for (const record of targetRecords) {
      if (record.pollMessageId && record.pollStatus !== "FINAL") {
        try {
          const channel = await textChannel(record.pollChannelId, "面接投票");
          const message = await channel.messages.fetch({ message: record.pollMessageId, force: true });
          let fetchedVotes = await fetchPollVoteCounts(message);
          let totalVotes = fetchedVotes.passVotes + fetchedVotes.failVotes;
          const expiresTimestamp = message.poll?.expiresTimestamp || 0;
          const expired = expiresTimestamp > 0 && expiresTimestamp <= Date.now();
          let finalized = Boolean(message.poll?.resultsFinalized) || expired;
          let processResult = finalized ? "面接投票を確定" : "面接投票結果を自動読込中";
          if (!finalized && totalVotes >= setting.pollVoteLimit) {
            const endedMessage = await message.poll.end();
            fetchedVotes = await fetchPollVoteCounts(endedMessage);
            totalVotes = fetchedVotes.passVotes + fetchedVotes.failVotes;
            finalized = true;
            processResult = `締切投票数${setting.pollVoteLimit}票に到達して確定`;
          }
          const verdict = pollVerdict(fetchedVotes.passVotes, totalVotes);
          const pollStatus = finalized ? "FINAL" : "PREVIEW";
          if (record.passVotes !== fetchedVotes.passVotes
            || record.failVotes !== fetchedVotes.failVotes
            || record.totalVotes !== totalVotes
            || record.pollStatus !== pollStatus) {
            await updateInterviewRecord(record, {
              N: pollStatus,
              O: pollStatus,
              P: fetchedVotes.passVotes,
              Q: fetchedVotes.failVotes,
              R: totalVotes,
              S: verdict,
              X: processResult,
            });
            Object.assign(record, {
              interviewStatus: pollStatus,
              pollStatus,
              passVotes: fetchedVotes.passVotes,
              failVotes: fetchedVotes.failVotes,
              totalVotes,
              verdict,
              processResult,
            });
            updatedCount += 1;
          }
        } catch (error) {
          await updateInterviewRecord(record, { X: `面接投票読込エラー: ${error.message}` });
        }
      }
      if (record.pollStatus === "FINAL" && ["合格", "不合格"].includes(record.verdict)) {
        const recruitmentSetting = recruitmentSettings.find((item) => item.roundName === record.roundName);
        await reconcileInterviewDecision(guild, setting, recruitmentSetting, record);
      }
      if (record.pollStatus === "FINAL" && record.verdict === "合格" && !record.passAnnouncementMessageId) {
        if (!setting.passChannelId) {
          await updateInterviewRecord(record, { X: `${record.processResult || "面接合格"} / 合格発表チャンネル待ち` });
          continue;
        }
        try {
          const announcement = await sendInterviewPass(setting, record);
          const result = `${record.processResult || "面接結果を確定"} / 面接合格発表済${announcement.applicantMatched ? "" : "（本人メンション未解決）"}`;
          await updateInterviewRecord(record, { X: result, Y: discordIdCell(announcement.messageId) });
          record.passAnnouncementMessageId = announcement.messageId;
          // Keep the per-applicant resend button available after success.
          try {
            await setAnnouncementRetryButton(record.pollChannelId, record.pollMessageId, "interview", record.id, true);
            announcementRetryButtonsInstalled.add(`iv:${record.id}`);
          } catch (buttonError) {
            console.warn(`面接合格発表の再送ボタン設置失敗 (${record.id}):`, buttonError.message);
          }
          announcedCount += 1;
        } catch (error) {
          await setAnnouncementRetryButton(record.pollChannelId, record.pollMessageId, "interview", record.id, true).catch((buttonError) => {
            console.warn(`面接合格発表の再送信ボタン設置失敗 (${record.id}):`, buttonError.message);
          });
          await updateInterviewRecord(record, { X: `面接合格発表エラー: ${error.message}` });
        }
      } else if (record.pollStatus === "FINAL"
        && record.verdict === "合格"
        && setting.passChannelId
        && record.passAnnouncementMessageId
        && !announcementRetryButtonsInstalled.has(`iv:${record.id}`)) {
        try {
          await setAnnouncementRetryButton(record.pollChannelId, record.pollMessageId, "interview", record.id, true);
          announcementRetryButtonsInstalled.add(`iv:${record.id}`);
        } catch (error) {
          console.warn(`面接合格発表の再送ボタン更新失敗 (${record.id}):`, error.message);
        }
      }
    }
    const activeCount = targetRecords.filter((record) => ["IN_PROGRESS", "READY", "PREVIEW"].includes(record.interviewStatus)).length;
    const finalCount = targetRecords.filter((record) => record.pollStatus === "FINAL").length;
    await writeInterviewStatus(
      setting,
      `STEP2稼働中: 面接${targetRecords.length}件 / 進行中${activeCount}件 / FINAL${finalCount}件 / 更新${updatedCount}件 / 合格発表${announcedCount}件`,
    );
  }
}

async function registerInterviewCommand(guild) {
  const commandData = new SlashCommandBuilder()
    .setName("mensetu")
    .setDescription("STEP1書類合格者から面接対象者を選択します")
    .setDMPermission(false)
    .addStringOption((option) => option
      .setName("search")
      .setDescription("応募ID・街での名前・Discordユーザー名で絞り込み")
      .setRequired(false))
    .toJSON();
  const commands = await guild.commands.fetch();
  const current = commands.find((command) => command.name === "mensetu");
  if (current) await current.edit(commandData);
  else await guild.commands.create(commandData);
  console.log("面接コマンド登録完了: /mensetu");
}

async function registerStaffProfileCommand(guild) {
  const commandData = new SlashCommandBuilder()
    .setName("syoin")
    .setDescription("DiscordユーザーIDから署員個票を表示します")
    .setDMPermission(false)
    .addUserOption((option) => option
      .setName("user")
      .setDescription("検索するDiscordユーザー")
      .setRequired(false))
    .addStringOption((option) => option
      .setName("id")
      .setDescription("17〜20桁のDiscordユーザーID（DC-付きでも可）")
      .setRequired(false))
    .toJSON();
  const commands = await guild.commands.fetch();
  const current = commands.find((command) => command.name === "syoin");
  if (current) await current.edit(commandData);
  else await guild.commands.create(commandData);
  console.log("署員個票コマンド登録完了: /syoin");
}

async function registerRankCommand(guild) {
  const commandData = new SlashCommandBuilder()
    .setName("rank")
    .setDescription("署員のランク変更を申請・実行します")
    .setDMPermission(false)
    .addStringOption((option) => option
      .setName("search")
      .setDescription("署員名またはDiscord IDで検索（省略すると一覧）")
      .setRequired(false))
    .toJSON();
  const commands = await guild.commands.fetch();
  const current = commands.find((command) => command.name === "rank");
  if (current) await current.edit(commandData);
  else await guild.commands.create(commandData);
  console.log("ランク操作コマンド登録完了: /rank");
}

async function rankEmployeeCandidates(search = "") {
  const employeeSheet = await readEmployees();
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const nameColumn = employeeSheet.headerMap.get("表示名");
  const rankColumn = employeeSheet.headerMap.get("適用ランク");
  const needle = String(search || "").trim().toLocaleLowerCase();
  return employeeSheet.rows
    .map((row) => ({
      id: normalizedDiscordId(row, employeeSheet),
      name: String(row[nameColumn] || "").trim(),
      rank: String(row[rankColumn] || "？？？？").trim() || "？？？？",
      employeeId: String(row[idColumn] || "").trim(),
    }))
    .filter((employee) => employee.id && employee.name && employee.rank !== "解雇者")
    .filter((employee) => !needle || [employee.id, employee.employeeId, employee.name].some((value) => value.toLocaleLowerCase().includes(needle)))
    .slice(0, 25);
}

function rankTargetMenu(candidates) {
  return new ActionRowBuilder().addComponents(
    new StringSelectMenuBuilder()
      .setCustomId("rank:target")
      .setPlaceholder("対象署員を選択")
      .addOptions(candidates.map((employee) => new StringSelectMenuOptionBuilder()
        .setLabel(truncateDiscord(employee.name, 100))
        .setDescription(`${employee.rank} / ${employee.id}`.slice(0, 100))
        .setValue(employee.id))),
  );
}

async function handleRankCommand(interaction) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  const settings = await readRankOperationSettings();
  if (!rankOperationPermission(interaction, settings)) {
    await interaction.editReply("ランク操作の管理ロールが設定されていないか、実行権限がありません。");
    return;
  }
  const candidates = await rankEmployeeCandidates(interaction.options?.getString?.("search") || "");
  if (!candidates.length) {
    await interaction.editReply("該当する在籍署員がいません。検索条件を変えて再実行してください。");
    return;
  }
  await interaction.editReply({
    content: "対象署員を選択してください。解雇者は対象外です。",
    components: [rankTargetMenu(candidates)],
  });
}

async function handleRankTargetSelect(interaction) {
  const discordId = interaction.values[0];
  await interaction.update({
    content: `対象Discord ID: **${discordId}**\n最初に操作項目を選択してください。`,
    components: [new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId(`rank:action:${discordId}`)
        .setPlaceholder("昇格・降格・警告・報告から選択")
        .addOptions([
          ["昇格", "ランクを上げる"],
          ["降格", "ランクを下げる"],
          ["警告", "ランク変更なしで警告を記録"],
          ["報告", "ランク変更なしで報告を記録"],
        ].map(([value, description]) => new StringSelectMenuOptionBuilder()
          .setLabel(value)
          .setDescription(description)
          .setValue(value))),
    )],
  });
}

async function handleRankActionSelect(interaction, discordId) {
  const action = interaction.values[0];
  if (["警告", "報告"].includes(action)) {
    await interaction.update({
      content: `操作: **${action}**\nランクは変更しません。備考を入力してください。`,
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`rank:reason:${discordId}:${encodeURIComponent(action)}:none`)
          .setLabel("備考を入力して実行")
          .setStyle(ButtonStyle.Primary),
      )],
    });
    return;
  }
  const [guild, rankMap] = await Promise.all([
    client.guilds.cache.get(guildId) || client.guilds.fetch(guildId),
    readRankMap(),
  ]);
  const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId);
  const current = assessMember(member, rankMap).rankName || "？？？？";
  const currentRank = rankByName(rankMap, current);
  const ranks = sortedRanks(rankMap)
    .filter((rank) => !currentRank || (action === "昇格" ? rank.priority < currentRank.priority : rank.priority > currentRank.priority))
    .slice(0, 25);
  if (!ranks.length) throw new Error(`${action}できる変更先ランクがありません。`);
  await interaction.update({
    content: `対象: **${truncateDiscord(member.displayName, 100)}**\n現在のランク: **${current}**\n操作: **${action}**\n変更後ランクを選択してください。`,
    components: [new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId(`rank:next:${discordId}:${encodeURIComponent(action)}`)
        .setPlaceholder("変更後ランクを選択")
        .addOptions(ranks.map((rank) => new StringSelectMenuOptionBuilder()
          .setLabel(truncateDiscord(rank.rankName, 100))
          .setDescription(`ロール: ${rank.roleName}`.slice(0, 100))
          .setValue(rank.rankName))),
    )],
  });
}

async function handleRankNextSelect(interaction, discordId, encodedAction) {
  const nextRank = interaction.values[0];
  const action = decodeURIComponent(encodedAction);
  await interaction.update({
    content: `操作: **${action}**\n変更後ランク: **${truncateDiscord(nextRank, 100)}**\n備考を入力して実行してください。備考は必須です。`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(`rank:reason:${discordId}:${encodeURIComponent(action)}:${encodeURIComponent(nextRank)}`)
        .setLabel("備考を入力して実行")
        .setStyle(ButtonStyle.Primary),
    )],
  });
}

async function handleRankReasonButton(interaction) {
  const [, , discordId, encodedAction, encodedRank] = interaction.customId.split(":");
  const modal = new ModalBuilder()
    .setCustomId(`rank:submit:${discordId}:${encodedAction}:${encodedRank}`)
    .setTitle("操作備考");
  const reason = new TextInputBuilder()
    .setCustomId("reason")
    .setLabel("備考（必須）")
    .setStyle(TextInputStyle.Paragraph)
    .setRequired(true)
    .setMaxLength(500)
    .setPlaceholder("昇格・降格・警告・報告の備考を入力してください");
  modal.addComponents(new ActionRowBuilder().addComponents(reason));
  await interaction.showModal(modal);
}

async function handleRankSubmit(interaction) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  const settings = await readRankOperationSettings();
  if (!rankOperationPermission(interaction, settings)) throw new Error("ランク操作の管理ロールが設定されていないか、実行権限がありません。");
  const [, , discordId, encodedAction, encodedRank] = interaction.customId.split(":");
  const action = decodeURIComponent(encodedAction);
  const nextRank = decodeURIComponent(encodedRank);
  const reason = String(interaction.fields.getTextInputValue("reason") || "").trim();
  if (!reason) throw new Error("備考は必須です。");
  const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);
  const [rankMap, employeeSheet] = await Promise.all([readRankMap(), readEmployees()]);
  const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId);
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const rowIndex = employeeSheet.rows.findIndex((row) => String(row[idColumn] || "").trim() === employeeId(discordId));
  if (rowIndex < 0) throw new Error("対象者が署員一覧に存在しません。最新の一覧を開き直してください。");
  const currentRank = assessMember(member, rankMap).rankName || "？？？？";
  const isRankMovement = ["昇格", "降格"].includes(action);
  if (isRankMovement && !nextRank) throw new Error("変更後ランクを選択してください。");
  const type = action;
  const threadId = action === "昇格"
    ? settings.promotionThreadId
    : action === "降格"
      ? settings.demotionThreadId
      : action === "警告"
        ? settings.warningThreadId
        : settings.reportThreadId;
  if (!/^\d{17,20}$/.test(threadId)) throw new Error(`${type}報告スレッドIDが未設定です。設定シートを確認してください。`);
  const thread = await client.channels.fetch(threadId);
  if (!thread?.isThread?.()) throw new Error(`${type}報告先がDiscordスレッドではありません。`);
  const result = isRankMovement
    ? await applySelectedRank(guild, member, rankMap, nextRank, `Discord /rank: ${reason}`)
    : { transition: `${currentRank}（変更なし）` };
  const rowNumber = rowIndex + 3;
  await writeActionResult(employeeSheet, rowNumber, {
    ...(isRankMovement ? {
      "適用ランク": nextRank,
      "Discordロール": rankByName(rankMap, nextRank)?.roleName || nextRank,
    } : {}),
    "操作結果": `完了: ${type}${isRankMovement ? ` ${result.transition}` : "（ランク変更なし）"} / 備考: ${reason}`,
    "操作日時": new Date().toISOString(),
  });
  const report = await thread.send({
    content: `【${type}報告】\n対象: <@${discordId}>\n変更: ${isRankMovement ? `${currentRank} → ${nextRank}` : "なし"}\n備考: ${reason}\n実行者: <@${interaction.user.id}>`,
    allowedMentions: { users: [discordId, interaction.user.id] },
  });
  const threadUrl = `https://discord.com/channels/${guildId}/${thread.id}/${report.id}`;
  await appendStaffReport({
    at: sheetDateTime(new Date()), discordId, name: member.displayName,
    actorId: interaction.user.id, actorName: interaction.user.username,
    previousRank: currentRank, nextRank: isRankMovement ? nextRank : "変更なし", reason, type, threadUrl,
    result: `完了: ${type}${isRankMovement ? ` ${result.transition}` : "（ランク変更なし）"}`,
  });
  await interaction.editReply({ content: `${type}を実行しました。\n${threadUrl}` });
}

function canViewStaffProfile(interaction) {
  return Boolean(
    interaction.memberPermissions?.has(PermissionFlagsBits.ManageRoles)
    || [...rosterRoleIds].some((id) => interaction.member?.roles?.cache?.has(id)),
  );
}

async function handleStaffProfileCommand(interaction) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  if (!canViewStaffProfile(interaction)) {
    await interaction.editReply("署員個票はPolice Officerロールを持つ署員または管理者のみ検索できます。");
    return;
  }
  const selectedUser = interaction.options?.getUser?.("user") || null;
  const rawId = String(interaction.options?.getString?.("id") || "").trim();
  if (selectedUser && rawId) {
    await interaction.editReply("user または id のどちらか一方だけを指定してください。");
    return;
  }
  const discordId = selectedUser?.id
    || rawId.replace(/^DC-/i, "").replace(/[<@!>]/g, "")
    || interaction.user.id;
  if (!/^\d{17,20}$/.test(discordId)) {
    await interaction.editReply("DiscordユーザーIDは17〜20桁の数字で入力してください。");
    return;
  }
  const [profile, guild] = await Promise.all([
    staffProfile(discordId),
    client.guilds.cache.get(guildId) || client.guilds.fetch(guildId),
  ]);
  if (!profile) {
    await interaction.editReply(`ID \`${discordId}\` は現在の署員一覧に見つかりません。Discord IDが違います、または署員一覧が未同期です。`);
    return;
  }
  const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId).catch(() => null);
  if (!member) {
    await interaction.editReply(`名簿には登録されていますが、Discordサーバー上で ID \`${discordId}\` を確認できません。Discord IDが違います。`);
    return;
  }
  await interaction.editReply({ embeds: [staffProfileEmbed(profile, member)] });
}

async function handleStaffProfileBoardCommand(interaction) {
  await interaction.deferReply({ flags: MessageFlags.Ephemeral });
  if (!canViewStaffProfile(interaction)) {
    await interaction.editReply("署員個票はPolice Officerロールを持つ署員または管理者のみ検索できます。");
    return;
  }
  const candidates = await rankEmployeeCandidates("");
  if (!candidates.length) {
    await interaction.editReply("現在、選択できる在籍署員がいません。");
    return;
  }
  const select = new StringSelectMenuBuilder()
    .setCustomId("command:syoin:target")
    .setPlaceholder("個票を表示する署員を選択")
    .addOptions(candidates.slice(0, 25).map((employee) => new StringSelectMenuOptionBuilder()
      .setLabel(truncateDiscord(employee.name, 100))
      .setDescription(`${employee.rank} / ${employee.id}`.slice(0, 100))
      .setValue(employee.id)));
  await interaction.editReply({
    content: candidates.length > 25 ? "先頭25名を表示しています。詳細検索は /syoin を使用してください。" : "個票を表示する署員を選択してください。",
    components: [new ActionRowBuilder().addComponents(select)],
  });
}

async function handleStaffProfileBoardTarget(interaction) {
  if (!canViewStaffProfile(interaction)) throw new Error("署員個票を閲覧する権限がありません。");
  const discordId = interaction.values[0];
  const [profile, guild] = await Promise.all([
    staffProfile(discordId),
    client.guilds.cache.get(guildId) || client.guilds.fetch(guildId),
  ]);
  if (!profile) throw new Error(`ID \`${discordId}\` は署員一覧に見つかりません。Discord IDが違います。`);
  const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId).catch(() => null);
  if (!member) throw new Error(`Discord ID \`${discordId}\` を確認できません。Discord IDが違います。`);
  await interaction.editReply({ content: "", embeds: [staffProfileEmbed(profile, member)], components: [] });
}

async function enqueueInterviewInteraction(label, interaction, work) {
  await enqueueDiscordOperation(label, async () => {
    try {
      await work();
    } catch (error) {
      const quotaLimited = registerSheetsQuotaError(error);
      const raw = String(error.message || error);
      const message = quotaLimited
        ? `Google Sheetsが一時的に混雑しています。約${sheetsBackoffRemainingSeconds()}秒後に自動復帰します。入力内容は消さず、同じ操作をもう一度行ってください。`
        : /すでに別の面接官|現在面接対象ではありません/.test(raw)
          ? `${raw}\n/mensetu で最新の候補一覧を開き直してください。`
          : `面接処理エラー: ${raw}\n画面を閉じずに再試行してください。`;
      console.error(message);
      if (interaction.deferred || interaction.replied) {
        await interaction.followUp({ content: message, flags: MessageFlags.Ephemeral }).catch(() =>
          interaction.editReply({ content: message }).catch(() => {}));
      } else {
        await interaction.reply({ content: message, flags: MessageFlags.Ephemeral }).catch(() => {});
      }
    }
  });
}

client.on("interactionCreate", async (interaction) => {
  if (interaction.guildId !== guildId) return;
  try {
    if (interaction.isButton() && interaction.customId.startsWith("command:")) {
      const action = interaction.customId.slice("command:".length);
      const expectedChannelId = commandBoardChannels.get(action === "mensetu" ? "mensetu" : "rank");
      if (!expectedChannelId || interaction.channelId !== expectedChannelId) {
        await interaction.reply({ content: "このボタンは別の専用チャンネルでのみ使用できます。", flags: MessageFlags.Ephemeral });
        return;
      }
      const remaining = consumeCommandBoardCooldown(interaction);
      if (remaining > 0) {
        await interaction.reply({ content: `連打防止中です。${remaining}秒後にもう一度押してください。`, flags: MessageFlags.Ephemeral });
        return;
      }
      const lockKey = `${interaction.user.id}:${interaction.customId}`;
      if (commandBoardLocks.has(lockKey)) {
        await interaction.reply({ content: "現在処理中です。完了するまでお待ちください。", flags: MessageFlags.Ephemeral });
        return;
      }
      commandBoardLocks.add(lockKey);
      try {
        if (action === "mensetu") await handleInterviewCommand(interaction);
        else if (action === "rank") await handleRankCommand(interaction);
        else if (action === "syoin") await handleStaffProfileBoardCommand(interaction);
        else await interaction.reply({ content: "この操作ボードのボタンは無効です。管理者に確認してください。", flags: MessageFlags.Ephemeral });
      } finally {
        commandBoardLocks.delete(lockKey);
      }
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("resend:doc:")) {
      await handleDocumentAnnouncementResend(interaction, interaction.customId.slice("resend:doc:".length));
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("resend:iv:")) {
      await handleInterviewAnnouncementResend(interaction, interaction.customId.slice("resend:iv:".length));
      return;
    }
    if (interaction.isChatInputCommand() && interaction.commandName === "syoin") {
      await handleStaffProfileCommand(interaction);
      return;
    }
    if (interaction.isChatInputCommand() && interaction.commandName === "rank") {
      await handleRankCommand(interaction);
      return;
    }
    if (interaction.isChatInputCommand() && interaction.commandName === "mensetu") {
      await handleInterviewCommand(interaction);
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId === "command:syoin:target") {
      if (!commandBoardChannels.get("syoin") || interaction.channelId !== commandBoardChannels.get("syoin")) {
        await interaction.reply({ content: "この個票選択はランク管理ボードの専用チャンネルでのみ使用できます。", flags: MessageFlags.Ephemeral });
        return;
      }
      await interaction.deferUpdate();
      await handleStaffProfileBoardTarget(interaction);
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId === "rank:target") {
      await handleRankTargetSelect(interaction);
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId.startsWith("rank:action:")) {
      await handleRankActionSelect(interaction, interaction.customId.slice("rank:action:".length));
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId.startsWith("rank:next:")) {
      const [, , discordId, encodedAction] = interaction.customId.split(":");
      await handleRankNextSelect(interaction, discordId, encodedAction);
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("rank:reason:")) {
      await handleRankReasonButton(interaction);
      return;
    }
    if (interaction.isModalSubmit() && interaction.customId.startsWith("rank:submit:")) {
      await handleRankSubmit(interaction);
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId.startsWith("iv:select:")) {
      await interaction.deferUpdate();
      await enqueueInterviewInteraction("面接対象者選択", interaction, () => handleInterviewSelect(interaction));
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("iv:page:")) {
      const [, , interviewId, questionIndex] = interaction.customId.split(":");
      await interaction.deferUpdate();
      await handleInterviewPageButton(interaction, interviewId, Number(questionIndex));
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId.startsWith("iv:jump:")) {
      const interviewId = interaction.customId.slice("iv:jump:".length);
      await interaction.deferUpdate();
      await handleInterviewJump(interaction, interviewId, Number(interaction.values[0]));
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("iv:q:")) {
      const [, , interviewId, questionIndex] = interaction.customId.split(":");
      await handleInterviewQuestionButton(interaction, interviewId, Number(questionIndex));
      return;
    }
    if (interaction.isModalSubmit() && interaction.customId.startsWith("iv:answer:")) {
      const [, , interviewId, questionIndex] = interaction.customId.split(":");
      await interaction.deferUpdate();
      await enqueueInterviewInteraction("面接回答保存", interaction, () => handleInterviewAnswerModal(interaction, interviewId, Number(questionIndex)));
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("iv:vote:")) {
      await interaction.deferUpdate();
      await enqueueInterviewInteraction("面接投票開始", interaction, () => handleInterviewVoteButton(interaction, interaction.customId.slice("iv:vote:".length)));
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith("iv:cancel:")) {
      await interaction.deferUpdate();
      await enqueueInterviewInteraction("面接キャンセル", interaction, () => handleInterviewCancelButton(interaction, interaction.customId.slice("iv:cancel:".length)));
    }
  } catch (error) {
    const quotaLimited = registerSheetsQuotaError(error);
    const isStaffProfileCommand = interaction.isChatInputCommand?.() && interaction.commandName === "syoin";
    const isRankCommand = interaction.isChatInputCommand?.() && interaction.commandName === "rank";
    const isCommandBoard = interaction.isButton?.() && interaction.customId?.startsWith("command:");
    const message = quotaLimited
      ? `Google Sheetsの読込制限が解除される約${sheetsBackoffRemainingSeconds()}秒後に、もう一度実行してください。`
      : `${isCommandBoard ? "操作ボード" : isStaffProfileCommand ? "署員個票" : isRankCommand ? "ランク操作" : "面接"}処理エラー: ${formatDiscordIdError(error)}`;
    console.error(message);
    if (interaction.deferred || interaction.replied) await interaction.editReply({ content: message, embeds: [], components: [] }).catch(() => {});
    else await interaction.reply({ content: message, flags: MessageFlags.Ephemeral }).catch(() => {});
  }
});

async function processTerminations() {
  const [terminationSheet, employeeSheet] = await Promise.all([readTerminations(), readEmployees()]);
  const completeColumn = terminationSheet.headerMap.get("手続き完了");
  const completedAtColumn = terminationSheet.headerMap.get("完了日");
  const deletionAtColumn = terminationSheet.headerMap.get("名簿削除予定日");
  const noteColumn = terminationSheet.headerMap.get("備考");
  const terminationEmployeeIdColumn = terminationSheet.headerMap.get("社員ID");
  const terminationDiscordIdColumn = terminationSheet.headerMap.get("DiscordユーザーID");
  const employeeIdColumn = employeeSheet.headerMap.get("社員ID");
  const now = new Date();
  const terminationUpdates = [];
  const employeeRangesToClear = [];
  const terminationRangesToClear = [];
  const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);

  for (let index = 0; index < terminationSheet.rows.length; index += 1) {
    const row = terminationSheet.rows[index];
    if (!String(row[terminationEmployeeIdColumn] || "").trim()) continue;
    if (!isChecked(row[completeColumn])) continue;

    const rowNumber = index + 3;
    let completedAt = parseSheetDate(row[completedAtColumn]);
    let deletionAt = parseSheetDate(row[deletionAtColumn]);
    if (!completedAt) {
      completedAt = now;
      deletionAt = new Date(completedAt.getTime() + retentionPeriodMs);
      terminationUpdates.push(...terminationRowUpdate(terminationSheet, rowNumber, {
        "完了日": sheetDateTime(completedAt),
        "名簿削除予定日": sheetDateTime(deletionAt),
      }));
      console.log(`退職手続き完了日・削除予定日を記録: ${row[terminationEmployeeIdColumn]}`);
      continue;
    }
    if (!deletionAt) {
      deletionAt = new Date(completedAt.getTime() + retentionPeriodMs);
      terminationUpdates.push(...terminationRowUpdate(terminationSheet, rowNumber, {
        "名簿削除予定日": sheetDateTime(deletionAt),
      }));
    }
    if (now.getTime() < deletionAt.getTime()) continue;

    const discordId = String(row[terminationDiscordIdColumn] || "").replace(/^'/, "").trim();
    const terminatedEmployeeId = String(row[terminationEmployeeIdColumn] || "").trim();
    try {
      if (!discordId) throw new Error("DiscordユーザーIDが空です");
      let member = null;
      try {
        member = await guild.members.fetch(discordId);
      } catch (error) {
        if (error.code !== 10007) throw error;
      }
      if (member) {
        const removedRoles = await removeAllEditableRoles(
          guild,
          member,
          "退職手続き完了から7日経過したため全ロールを自動解除",
        );
        console.log(`解雇者の全ロール解除: ${member.displayName} (${removedRoles.join(", ") || "解除対象なし"})`);
      }
    } catch (error) {
      const message = error.response?.data?.message || error.message || String(error);
      const currentNote = String(row[noteColumn] || "").trim();
      const errorNote = `自動削除エラー: ${message}`;
      terminationUpdates.push(...terminationRowUpdate(terminationSheet, rowNumber, {
        "備考": currentNote.includes(errorNote) ? currentNote : [currentNote, errorNote].filter(Boolean).join(" / "),
      }));
      console.error(`解雇者の自動削除失敗: ${terminatedEmployeeId}`, message);
      continue;
    }

    const employeeIndex = employeeSheet.rows.findIndex((employeeRow) => {
      const candidateDiscordId = normalizedDiscordId(employeeRow, employeeSheet);
      return (discordId && candidateDiscordId === discordId)
        || String(employeeRow[employeeIdColumn] || "").trim() === terminatedEmployeeId;
    });
    if (employeeIndex >= 0) employeeRangesToClear.push(`'${employeeSheetName}'!A${employeeIndex + 3}:ZZ${employeeIndex + 3}`);
    terminationRangesToClear.push(`'${terminationSheetName}'!A${rowNumber}:I${rowNumber}`);
    terminationRangesToClear.push(`'${terminationSheetName}'!K${rowNumber}:K${rowNumber}`);
    console.log(`解雇者記録を自動削除: ${terminatedEmployeeId}`);
  }

  if (terminationUpdates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data: terminationUpdates },
    });
  }
  if (employeeRangesToClear.length) {
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId,
      requestBody: { ranges: [...new Set(employeeRangesToClear)] },
    });
    await sortEmployees(employeeSheet);
  }
  if (terminationRangesToClear.length) {
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId,
      requestBody: { ranges: [...new Set(terminationRangesToClear)] },
    });
  }
}

async function fullSync() {
  const guild = await client.guilds.fetch(guildId);
  const members = await guild.members.fetch();
  const rankMap = await readRankMap();
  const dismissal = dismissedRank(rankMap);
  const missingRosterRoles = [...rosterRoleIds].filter((id) => !guild.roles.cache.has(id));
  if (missingRosterRoles.length) {
    throw new Error(`ROSTER_ROLE_IDがこのサーバーに存在しません: ${missingRosterRoles.join(", ")}`);
  }
  if (!dismissal) throw new Error("ランク設定に有効な「解雇者」ロールがありません");
  if (!guild.roles.cache.has(dismissal.roleId)) {
    throw new Error(`解雇者ロールがこのサーバーに存在しません: ${dismissal.roleId}`);
  }
  const rosterNames = [...rosterRoleIds].map((id) => guild.roles.cache.get(id)?.name || id);
  const eligibleCount = members.filter((member) =>
    !member.user.bot &&
    [...rosterRoleIds].some((id) => member.roles.cache.has(id)) &&
    ![...excludeRoleIds].some((id) => member.roles.cache.has(id)) &&
    !member.roles.cache.has(dismissal.roleId),
  ).size;
  console.log(`名簿対象ロール: ${rosterNames.join(", ")}`);
  console.log(`解雇者判定ロール: ${dismissal.roleName} (${dismissal.roleId})`);
  console.log(`名簿対象人数: ${eligibleCount}人`);
  if (eligibleCount === 0) {
    throw new Error("名簿対象者が0人です。ROSTER_ROLE_IDが全署員に共通するロールか確認してください。");
  }
  await consolidateEmployeeDuplicates();
  const employeeSheet = await readEmployees();
  const context = { rankMap, employeeSheet, pendingData: [], pendingClearRanges: [], silent: true };
  for (const member of members.values()) await syncMember(member, context);
  if (context.pendingClearRanges.length) {
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId,
      requestBody: { ranges: [...new Set(context.pendingClearRanges)] },
    });
  }
  if (context.pendingData.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data: context.pendingData },
    });
  }
  await sortEmployees(employeeSheet);
  console.log(`全件同期完了: ${members.size}人確認`);
}

async function markRemoved(member) {
  if (member.guild.id !== guildId || member.user.bot) return;
  lastSynchronizedRanks.delete(member.id);
  const employeeSheet = await readEmployees();
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const index = employeeSheet.rows.findIndex((row) => String(row[idColumn] || "") === employeeId(member.id));
  if (index < 0) return;
  const rowNumber = index + 3;
  await clearEmployeeRow(employeeSheet, rowNumber);
  console.log(`Discordサーバー脱退のため名簿削除: ${member.displayName || member.user.username}`);
}

client.once("clientReady", async () => {
  console.log(`Discord接続完了: ${client.user.tag}`);
  // These are repair operations, not runtime prerequisites.  Running them on
  // every Railway restart can exhaust the Sheets write quota before form sync.
  if (process.env.REPAIR_SETTINGS_UI === "1") {
    try {
      await ensureUnifiedSettingsCategoryValidation();
    } catch (error) {
      console.error("設定カテゴリの入力規則更新失敗:", error.message);
    }
  }
  try {
    await auditUnifiedSettings(true);
  } catch (error) {
    console.error("統合設定の起動時監査失敗:", error.message);
  }
  if (process.env.REPAIR_SETTINGS_UI === "1") {
    try {
      await ensureRankAndEmployeeValidation();
    } catch (error) {
      console.error("ランク設定・署員一覧の入力規則更新失敗:", error.message);
    }
  }
  try {
    const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);
    await validateRankReportDestinations(guild);
  } catch (error) {
    console.error("ランク報告先の自動準備失敗:", error.message);
  }
  try {
    const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);
    await registerInterviewCommand(guild);
    await registerStaffProfileCommand(guild);
    await registerRankCommand(guild);
    const boardSettings = await readCommandBoardSettings();
    await ensureCommandBoard(guild, "interview", boardSettings.interviewChannelId, boardSettings.interviewMessageId, boardSettings.interviewEnabled);
    await ensureCommandBoard(guild, "rank", boardSettings.rankChannelId, boardSettings.rankMessageId, boardSettings.rankEnabled);
  } catch (error) {
    console.error("Discordコマンド登録失敗:", error.message);
  }
  const actionPollInterval = Math.max(Number(process.env.ACTION_POLL_INTERVAL_MS || 30000), 30000);
  const initialPollDelay = Math.max(actionPollInterval, 60000);
  const pollSheetActions = async () => {
    await enqueueSheets("シート操作・応募・退職処理", async () => {
      if (Date.now() < sheetsQuotaBackoffUntil) {
        console.log(`Google Sheets読込待機中: 残り約${sheetsBackoffRemainingSeconds()}秒`);
        return;
      }
      // Form import is the highest priority.  A quota failure in an unrelated
      // maintenance job must not prevent the recruitment importer from running.
      const jobs = [
        ["書類選考", processRecruitmentApplications],
        ["統合設定監査", auditUnifiedSettings],
        ["署員一覧操作", processSheetActions],
        ["ボーナス配布", processBonusDistribution],
        ["面接", processInterviewPolls],
        ["採用手続き", processOnboarding],
        ["退職手続き", processTerminations],
      ];
      let completedWithoutQuotaError = true;
      for (const [label, job] of jobs) {
        if (label !== "書類選考" && Date.now() < maintenanceQuotaBackoffUntil) continue;
        try {
          await job();
        } catch (error) {
          const message = error.response?.data?.error?.message || error.message || String(error);
          console.error(`[${label}]`, message);
          if (registerSheetsQuotaError(error)) {
            completedWithoutQuotaError = false;
            if (label !== "書類選考") {
              // Maintenance jobs can be write-heavy. Back them off locally so
              // they cannot starve the form importer in the next cycle.
              maintenanceQuotaBackoffUntil = Date.now() + 120000;
              sheetsQuotaBackoffUntil = Math.min(sheetsQuotaBackoffUntil, Date.now());
            }
            // Keep independent jobs isolated; recruitment remains first.
            continue;
          }
        }
      }
      if (completedWithoutQuotaError) sheetsQuotaBackoffLevel = 0;
    });
    setTimeout(pollSheetActions, actionPollInterval);
  };
  const pollSettingsView = async () => {
    if (Date.now() >= sheetsQuotaBackoffUntil) {
      try {
        await applyUnifiedSettingsView();
      } catch (error) {
        if (!registerSheetsQuotaError(error)) console.error("[統合設定表示]", error.message || error);
      }
    }
    setTimeout(pollSettingsView, 10000);
  };
  // Start form/application polling independently of the expensive member
  // backfill.  A large fullSync must never consume the Sheets write quota in
  // normal operation; run it only when explicitly requested during maintenance.
  setTimeout(pollSheetActions, initialPollDelay);
  setTimeout(pollSettingsView, 10000);
  if (process.env.FULL_SYNC_ON_START === "1") {
    setTimeout(() => enqueueSheets("起動時全件同期", fullSync), 120000);
  }
  console.log(`シート操作・応募・退職処理監視: ${actionPollInterval}ms間隔`);
  console.log(`起動後の初回シート確認: ${initialPollDelay}ms後`);
  console.log(`応募フォーム確認: ${recruitmentPollIntervalMs}ms間隔`);
  console.log(`面接投票確認: ${interviewPollIntervalMs}ms間隔`);
});
client.on("guildMemberAdd", (member) => enqueueSheets("加入", async () => {
  await syncMember(member);
  await sortEmployees(await readEmployees());
}));
client.on("guildMemberUpdate", (_oldMember, newMember) => enqueueSheets("ロール変更", async () => {
  await syncMember(newMember);
  await sortEmployees(await readEmployees());
}));
client.on("guildMemberRemove", (member) => enqueueSheets("脱退", () => markRemoved(member)));

const port = Number(process.env.PORT || 3000);
http.createServer((_req, res) => {
  res.writeHead(200, { "content-type": "text/plain; charset=utf-8" });
  res.end(client.isReady() ? "Discord Sheets Bot: OK" : "Discord Sheets Bot: starting");
}).listen(port, () => console.log(`Health check: ${port}`));

async function connectDiscord() {
  try {
    await client.login(process.env.DISCORD_BOT_TOKEN);
  } catch (error) {
    console.error("Discord接続失敗。30秒後に自動再接続します:", error.message);
    setTimeout(connectDiscord, 30000);
  }
}

connectDiscord();
