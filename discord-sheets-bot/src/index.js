import http from "node:http";
import { createHash } from "node:crypto";
import {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  Client,
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

const client = new Client({
  intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMembers, GatewayIntentBits.GuildMessages],
});

let queue = Promise.resolve();
const lastSynchronizedRanks = new Map();
function enqueue(label, work) {
  queue = queue.then(work).catch((error) => {
    const message = error.response?.data?.error?.message || error.message || error;
    console.error(`[${label}]`, message);
  });
  return queue;
}

async function readRankMap() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'ランク設定'!A3:F1000",
  });
  const map = new Map();
  for (const row of response.data.values || []) {
    const priority = Number(row[0]);
    const rankName = String(row[1] || "").trim();
    const roleId = String(row[2] || "").trim();
    const roleName = String(row[3] || rankName).trim();
    const enabled = String(row[5] || "") === "はい";
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
const botHeaders = ["社員ID", "表示名", "Discordロール", "適用ランク", bonusFactorHeader, rankSelectionHeader, actionTriggerHeader, "操作結果", "操作日時", sortPriorityHeader];
const terminationHeaders = ["社員ID", "表示名", "DiscordユーザーID", "最終ランク", "解雇日", "手続き完了", "対応署員", "完了日", "名簿削除予定日", "名簿削除状況", "備考"];
const terminationSheetName = "解雇者管理";
const employeeSheetId = 1100459512;
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
const onboardingHeaders = [
  "応募ID", "面接ID", "募集回", "受験者名", "Discordユーザー名", "DiscordユーザーID",
  "面接合格日時", "手続き完了", "対応署員", "完了日時", "採用状態", "ロール処理結果",
  "最終エラー", "呼出日時", "呼出者", "備考",
];
let lastInterviewPollAt = 0;
const interviewPollIntervalMs = Math.max(Number(process.env.INTERVIEW_POLL_INTERVAL_MS || 60000), 30000);

async function readEmployees() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'従業員'!A2:ZZ1000",
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
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'従業員'!${columnLetter(start)}2:${columnLetter(end)}2`,
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
  if (index === undefined) throw new Error(`従業員シートに「${header}」列がありません。`);
  const col = columnLetter(index);
  return `'従業員'!${col}${rowNumber}`;
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
    [bonusFactorHeader]: `=IF(${id}="","",IFNA(XLOOKUP(${appliedRank},'ランク設定'!$B$3:$B$1000,'ランク設定'!$E$3:$E$1000),0))`,
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
    throw new Error("従業員シートの並び替え用見出しが見つかりません");
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
  const range = `'従業員'!A${rowNumber}:ZZ${rowNumber}`;
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
    console.log(`ロールなしのため名簿削除: ${member.displayName}`);
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
      console.log(`退職者の名簿対象ロールを自動解除: ${member.displayName} (${removedRoles.join(", ")})`);
    }
    const appliedRankColumn = employeeSheet.headerMap.get("適用ランク");
    const previousRank = assessed.rankName || (index >= 0
      ? String(employeeSheet.rows[index][appliedRankColumn] || "").trim()
      : "");
    await upsertTerminationRecord(member, previousRank);
    if (index >= 0) {
      await applyEmployeeUpdates(
        rowUpdate(employeeSheet, targetRow, {
          "表示名": member.displayName,
          "Discordロール": dismissal.roleName,
          "適用ランク": previousRank,
          [sortPriorityHeader]: 10000,
        }),
        context,
      );
    }
    lastSynchronizedRanks.set(member.id, "解雇");
    console.log(`解雇者ロール検知: ${member.displayName}${previousRank ? ` (最終ランク: ${previousRank})` : ""}`);
    return;
  }

  if (!hasRosterRole && assessed.rankName && index >= 0 && !hasExcludeRole) {
    await ensureRosterRole(member.guild, member, "階級ロールが残っているためPolice Officerを自動復元");
    hasRosterRole = true;
    console.log(`Police Officerロール自動復元: ${member.displayName}`);
  }

  const eligible = hasRosterRole && !hasExcludeRole;

  if (!eligible) {
    if (index < 0) return;
    await clearEmployeeRow(employeeSheet, targetRow, context);
    lastSynchronizedRanks.delete(member.id);
    console.log(`名簿対象外: ${member.displayName}`);
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
  console.log(`同期完了: ${member.displayName} → ${appliedRank}`);
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
    "基本ボーナス", bonusFactorHeader, legacyBonusFactorHeader, "調整額", "見込ボーナス", rankSelectionHeader, actionTriggerHeader, "操作結果", "操作日時", sortPriorityHeader,
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
      ranges: uniqueDuplicateIndexes.map((index) => `'従業員'!A${index + 3}:ZZ${index + 3}`),
    },
  });
  await sortEmployees(employeeSheet);
  console.log(`重複統合完了: ${uniqueDuplicateIndexes.length}行を統合`);
  return uniqueDuplicateIndexes.length;
}

async function writeActionResult(employeeSheet, rowNumber, fields) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: rowUpdate(employeeSheet, rowNumber, fields),
    },
  });
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
        "操作結果": "エラー: 社員IDからDiscordユーザーIDを判別できません",
        "操作日時": new Date().toISOString(),
      });
      continue;
    }
    if (!discordId) continue;

    try {
      const member = guild.members.cache.get(discordId) || await guild.members.fetch(discordId);
      const dismissal = dismissedRank(rankMap);
      const discordRank = dismissal && member.roles.cache.has(dismissal.roleId)
        ? "解雇"
        : assessMember(member, rankMap).rankName || "？？？？";

      if (triggered) {
        const selectedRank = String(row[targetColumn] || "").trim();
        await writeActionResult(employeeSheet, rowNumber, { "操作結果": `処理中: ${selectedRank || "未選択"}` });
        if (!selectedRank) throw new Error("変更後ランクを選択してください");
        const result = await applySelectedRank(guild, member, rankMap, selectedRank, "Google Sheetsの統合操作");
        await syncMember(await guild.members.fetch(discordId));
        needsSort = true;
        await writeActionResult(employeeSheet, rowNumber, {
          [rankSelectionHeader]: "",
          [actionTriggerHeader]: false,
          "操作結果": `完了: ${result.transition}`,
          "操作日時": new Date().toISOString(),
        });
        console.log(`シート操作完了: ${member.displayName} ${result.transition}`);
        continue;
      }

      if (discordRank === "解雇") {
        lastSynchronizedRanks.set(discordId, discordRank);
        continue;
      }

      const sheetRank = String(row[appliedRankColumn] || "").trim();
      if (!sheetRank || sheetRank === discordRank) {
        lastSynchronizedRanks.set(discordId, discordRank);
        continue;
      }

      const lastRank = lastSynchronizedRanks.get(discordId);
      if (lastRank === discordRank) {
        const result = await applySelectedRank(guild, member, rankMap, sheetRank, "Google Sheetsの適用ランク直接編集");
        await syncMember(await guild.members.fetch(discordId));
        needsSort = true;
        await writeActionResult(employeeSheet, rowNumber, {
          "操作結果": `完了: スプシ → Discord (${result.transition})`,
          "操作日時": new Date().toISOString(),
        });
        console.log(`双方向同期: スプシ → Discord ${member.displayName} ${result.transition}`);
      } else {
        await syncMember(member);
        needsSort = true;
        await writeActionResult(employeeSheet, rowNumber, {
          "操作結果": `完了: Discord → スプシ (${discordRank})`,
          "操作日時": new Date().toISOString(),
        });
        console.log(`双方向同期: Discord → スプシ ${member.displayName} ${discordRank}`);
      }
    } catch (error) {
      const message = error.response?.data?.message || error.message || String(error);
      await writeActionResult(employeeSheet, rowNumber, {
        [actionTriggerHeader]: false,
        "操作結果": `エラー: ${message}`,
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
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${bonusRoundSettingsSheetName}'!A2:C100`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const rows = response.data.values || [];
  let index = rows.findIndex((row) => String(row[0] || "").trim() === roundName);
  if (index < 0) index = rows.findIndex((row) => !String(row[0] || "").trim());
  if (index < 0) index = rows.length;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${bonusRoundSettingsSheetName}'!A${index + 2}:C${index + 2}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[roundName, poolAmount, sheetDateTime(loadedAt)]] },
  });
}

function bonusLedgerRow(rowNumber, roundName, employee, loadedAt, existingRow = null) {
  const basicPay = `=IF(B${rowNumber}="","",IF(SUMIF($L$12:$L$1000,$L${rowNumber},$E$12:$E$1000)=0,0,ROUNDDOWN(IFNA(XLOOKUP($L${rowNumber},'${bonusRoundSettingsSheetName}'!$A$2:$A$100,'${bonusRoundSettingsSheetName}'!$B$2:$B$100),0)*E${rowNumber}/SUMIF($L$12:$L$1000,$L${rowNumber},$E$12:$E$1000),-7)))`;
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
    await writeBonusTop({ N4: "処理中: 従業員を読込中" });
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
    if (!employees.length) throw new Error("従業員ページに読込対象がいません");

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

function recruitmentSettingFromRow(row, index) {
  const requestedDuration = Number(row[5]);
  const requestedVoteLimit = Number(row[6]);
  return {
    rowNumber: index + 4,
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
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${recruitmentSettingsSheetName}'!J${setting.rowNumber}:L${setting.rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[status, sheetDateTime(new Date()), false]] },
  });
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

async function sendApplicationPass(setting, application) {
  const channel = await textChannel(setting.passChannelId, "合格発表");
  const userId = await resolveApplicantDiscordUserId(application.discordId);
  const content = [setting.passMessage || "合格が決定しました。", userId ? `<@${userId}>` : ""]
    .filter(Boolean)
    .join("\n");
  const message = await channel.send({
    content,
    allowedMentions: { users: userId ? [userId] : [] },
    nonce: discordNonce("docpass", application.id),
    enforceNonce: true,
  });
  return { messageId: message.id, applicantMatched: Boolean(userId) };
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
            passAnnouncementCount += 1;
          } catch (error) {
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
          } catch (error) {
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `'${applicationSheetName}'!X${application.rowNumber}`,
              valueInputOption: "RAW",
              requestBody: { values: [[`合格発表更新エラー: ${error.message}`]] },
            });
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
      await writeRecruitmentStatus(setting, `エラー: ${message}`);
      console.error(`応募フォーム連携失敗: ${setting.roundName}`, message);
    }
  }
}

function interviewSettingFromRow(row, index) {
  const requestedDuration = Number(row[6]);
  const requestedVoteLimit = Number(row[7]);
  return {
    rowNumber: index + 4,
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
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${interviewSettingsSheetName}'!A4:P100`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return (response.data.values || [])
    .map(interviewSettingFromRow)
    .filter((setting) => setting.enabled);
}

async function writeInterviewStatus(setting, status) {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${interviewSettingsSheetName}'!L${setting.rowNumber}:M${setting.rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[status, sheetDateTime(new Date())]] },
  });
}

async function readInterviewQuestions(questionSet) {
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

async function readInterviewRecords() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${interviewManagementSheetName}'!A11:AD1000`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return (response.data.values || [])
    .map((row, index) => interviewRecordFromRow(row, index + 11))
    .filter((record) => record.id);
}

async function readStep1Applications() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${applicationSheetName}'!A11:AD1000`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return (response.data.values || [])
    .map((row, index) => applicationFromSheetRow(row, index + 11))
    .filter((application) => application.id);
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
    application.pollStatus === "FINAL"
    && application.verdict === "合格"
    && application.passAnnouncementMessageId
    && application.documentRoleStatus.startsWith("書類合格ロール付与済")
    && (!setting.roundName || application.roundName === setting.roundName)
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
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${interviewManagementSheetName}'!A11:A1000`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  const ids = response.data.values || [];
  const emptyIndex = Array.from({ length: 990 }, (_, index) => index)
    .find((index) => !String(ids[index]?.[0] || "").trim());
  if (emptyIndex === undefined) throw new Error("面接管理シートの保存行が不足しています。");
  const rowNumber = emptyIndex + 11;
  const id = nextInterviewId(application.id, records);
  const applicantDiscordId = await resolveApplicantDiscordUserId(application.discordId);
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
  return interviewRecordFromRow(values, rowNumber);
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
  return [new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`iv:page:${record.id}:${Math.max(safeIndex - 1, 0)}`)
      .setLabel("前の質問")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safeIndex === 0),
    new ButtonBuilder()
      .setCustomId(`iv:q:${record.id}:${safeIndex}`)
      .setLabel(current && answers[current.id] ? "この回答を修正" : "この質問に回答")
      .setStyle(ButtonStyle.Primary),
    new ButtonBuilder()
      .setCustomId(`iv:page:${record.id}:${Math.min(safeIndex + 1, Math.max(questions.length - 1, 0))}`)
      .setLabel("次の質問")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safeIndex >= questions.length - 1),
    new ButtonBuilder()
      .setCustomId(`iv:vote:${record.id}`)
      .setLabel("面接を完了して投票開始")
      .setStyle(ButtonStyle.Success)
      .setDisabled(!interviewReady(record)),
    new ButtonBuilder()
      .setCustomId(`iv:cancel:${record.id}`)
      .setLabel("面接をキャンセル")
      .setStyle(ButtonStyle.Danger),
  )];
}

function interviewSessionEmbed(record, questionIndex = 0) {
  const questions = interviewSnapshotQuestions(record);
  const answers = interviewAnswers(record);
  const safeIndex = Math.max(0, Math.min(questionIndex, Math.max(questions.length - 1, 0)));
  const answered = questions.filter((question) => String(answers[question.id] || "").trim()).length;
  const valueBudget = Math.max(140, Math.min(900, Math.floor(4700 / Math.max(questions.length, 1))));
  const fields = questions.slice(0, 24).map((question, index) => ({
    name: `${index === safeIndex ? "▶" : answers[question.id] ? "✅" : "⬜"} Q${index + 1}${question.required ? "【必須】" : "【任意】"}`,
    value: truncateDiscord(`${question.text}${question.note ? `\n補足: ${question.note}` : ""}${answers[question.id] ? `\n回答: ${answers[question.id]}` : ""}`, valueBudget),
  }));
  return new EmbedBuilder()
    .setColor(interviewReady(record) ? 0x16a34a : 0x2563eb)
    .setTitle(truncateDiscord(`STEP2 面接進行｜${record.applicantName || record.applicationId}`, 256))
    .setDescription(`応募ID: **${record.applicationId}**\n面接ID: **${record.id}**\n回答状況: **${answered}/${questions.length}問**`)
    .addFields(fields)
    .setFooter({ text: "質問は常時表示されます。前後ボタンで選び、1問ずつ回答してください。" });
}

async function interviewRecordById(id) {
  return (await readInterviewRecords()).find((record) => record.id === id) || null;
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
      && (!setting.roundName || record.roundName === setting.roundName)
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
  const search = interaction.options.getString("search") || "";
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
  const setting = settings.find((item) => !item.roundName || item.roundName === record.roundName);
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

async function sendInterviewPass(setting, record) {
  if (!setting.passChannelId) throw new Error("面接合格発表チャンネルIDが未設定です");
  const channel = await textChannel(setting.passChannelId, "面接合格発表");
  const userId = record.applicantDiscordId || await resolveApplicantDiscordUserId(record.applicantDiscordName);
  const content = [setting.passMessage, userId ? `<@${userId}>` : ""].filter(Boolean).join("\n");
  const message = await channel.send({
    content,
    allowedMentions: { users: userId ? [userId] : [] },
    nonce: discordNonce("ivpass", record.id),
    enforceNonce: true,
  });
  return { messageId: message.id, applicantMatched: Boolean(userId) };
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
    const recruitmentSetting = recruitmentSettings.find((item) => item.roundName === setting.roundName);
    const targetRecords = records.filter((record) => !setting.roundName || record.roundName === setting.roundName);
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
          announcedCount += 1;
        } catch (error) {
          await updateInterviewRecord(record, { X: `面接合格発表エラー: ${error.message}` });
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

async function enqueueInterviewInteraction(label, interaction, work) {
  await enqueue(label, async () => {
    try {
      await work();
    } catch (error) {
      const message = `面接処理エラー: ${error.message || error}`;
      console.error(message);
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ content: message, embeds: [], components: [] }).catch(() => {});
      } else {
        await interaction.reply({ content: message, flags: MessageFlags.Ephemeral }).catch(() => {});
      }
    }
  });
}

client.on("interactionCreate", async (interaction) => {
  if (interaction.guildId !== guildId) return;
  try {
    if (interaction.isChatInputCommand() && interaction.commandName === "mensetu") {
      await handleInterviewCommand(interaction);
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
      await enqueueInterviewInteraction("面接質問切替", interaction, () => handleInterviewPageButton(interaction, interviewId, Number(questionIndex)));
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
    const message = `面接処理エラー: ${error.message || error}`;
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
    if (employeeIndex >= 0) employeeRangesToClear.push(`'従業員'!A${employeeIndex + 3}:ZZ${employeeIndex + 3}`);
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
  const context = { rankMap, employeeSheet, pendingData: [], pendingClearRanges: [] };
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
  try {
    const guild = client.guilds.cache.get(guildId) || await client.guilds.fetch(guildId);
    await registerInterviewCommand(guild);
  } catch (error) {
    console.error("面接コマンド登録失敗:", error.message);
  }
  const actionPollInterval = Math.max(Number(process.env.ACTION_POLL_INTERVAL_MS || 10000), 5000);
  const initialPollDelay = Math.max(actionPollInterval, 60000);
  const pollSheetActions = async () => {
    await enqueue("シート操作・応募・退職処理", async () => {
      await processSheetActions();
      await processBonusDistribution();
      await processRecruitmentApplications();
      await processInterviewPolls();
      await processOnboarding();
      await processTerminations();
    });
    setTimeout(pollSheetActions, actionPollInterval);
  };
  enqueue("起動時全件同期", fullSync)
    .finally(() => setTimeout(pollSheetActions, initialPollDelay));
  console.log(`シート操作・応募・退職処理監視: ${actionPollInterval}ms間隔`);
  console.log(`起動後の初回シート確認: ${initialPollDelay}ms後`);
  console.log(`応募フォーム確認: ${recruitmentPollIntervalMs}ms間隔`);
  console.log(`面接投票確認: ${interviewPollIntervalMs}ms間隔`);
});
client.on("guildMemberAdd", (member) => enqueue("加入", async () => {
  await syncMember(member);
  await sortEmployees(await readEmployees());
}));
client.on("guildMemberUpdate", (_oldMember, newMember) => enqueue("ロール変更", async () => {
  await syncMember(newMember);
  await sortEmployees(await readEmployees());
}));
client.on("guildMemberRemove", (member) => enqueue("脱退", () => markRemoved(member)));

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
