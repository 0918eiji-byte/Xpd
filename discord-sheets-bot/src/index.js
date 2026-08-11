import http from "node:http";
import { Client, EmbedBuilder, GatewayIntentBits } from "discord.js";
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
const applicationSheetName = "応募管理";
const applicationSheetId = 2081134610;
let displayedApplicationRound = "";
let lastRecruitmentPollAt = 0;
const recruitmentPollIntervalMs = Math.max(Number(process.env.RECRUITMENT_POLL_INTERVAL_MS || 60000), 30000);

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
    manualTrigger: isChecked(row[9]),
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
    ranges: [`'${recruitmentSettingsSheetName}'!A4:K100`],
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
    range: `'${recruitmentSettingsSheetName}'!H${setting.rowNumber}:J${setting.rowNumber}`,
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
              endColumnIndex: 26,
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
  };
}

function truncateDiscord(value, maxLength = 360) {
  const text = String(value || "").trim() || "（未記入）";
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function applicationLink(rowNumber) {
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${applicationSheetId}&range=A${rowNumber}:X${rowNumber}`;
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

function pollStateFromMessage(message) {
  const poll = message.poll;
  if (!poll) throw new Error("Discordメッセージに投票データがありません");
  const answers = [...poll.answers.values()];
  const passVotes = answers.find((answer) => answer.text === "合格")?.voteCount || 0;
  const failVotes = answers.find((answer) => answer.text === "不合格")?.voteCount || 0;
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
  const message = await channel.messages.fetch(application.pollMessageId);
  const state = pollStateFromMessage(message);
  if (state.pollStatus === "PREVIEW" && state.totalVotes >= setting.pollVoteLimit) {
    const endedMessage = await message.poll.end();
    return {
      ...pollStateFromMessage(endedMessage),
      pollStatus: "FINAL",
      processResult: `締切投票数${setting.pollVoteLimit}票に到達して確定`,
    };
  }
  return state;
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
      range: `'${applicationSheetName}'!A11:Z1000`,
      valueRenderOption: "UNFORMATTED_VALUE",
    }),
  ]);
  const selectedRound = String(topResponse.data.values?.[0]?.[1] || "").trim();
  if (selectedRound) await applyApplicationRoundFilter(selectedRound);

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
              range: `'${applicationSheetName}'!A${rowNumber}:Z${rowNumber}`,
              values: [values],
            })),
          },
        });
      }

      let pollCreatedCount = 0;
      let pollUpdatedCount = 0;
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

        if (application.pollStatus === "FINAL") continue;
        try {
          const state = await fetchApplicationPoll(application, setting);
          if (!applicationPollStateChanged(application, state)) continue;
          await writeApplicationPollState(application.rowNumber, state);
          Object.assign(application, state);
          pollUpdatedCount += 1;
        } catch (error) {
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${applicationSheetName}'!X${application.rowNumber}`,
            valueInputOption: "RAW",
            requestBody: { values: [[`投票読込エラー: ${error.message}`]] },
          });
        }
      }

      const previewCount = roundApplications.filter((application) => application.pollStatus === "PREVIEW").length;
      const finalCount = roundApplications.filter((application) => application.pollStatus === "FINAL").length;
      const channelWait = !setting.pollChannelId ? " / 投票チャンネル待ち" : "";
      await writeRecruitmentStatus(
        setting,
        `稼働中: 応募${roundApplications.length}件 / 新規${newRows.length}件 / 投票作成${pollCreatedCount}件 / 更新${pollUpdatedCount}件 / PREVIEW${previewCount}件 / FINAL${finalCount}件 / 締切${setting.pollDurationHours}時間または${setting.pollVoteLimit}票${channelWait}`,
      );
    } catch (error) {
      const message = error.response?.data?.error?.message || error.message || String(error);
      await writeRecruitmentStatus(setting, `エラー: ${message}`);
      console.error(`応募フォーム連携失敗: ${setting.roundName}`, message);
    }
  }
}

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

client.once("clientReady", () => {
  console.log(`Discord接続完了: ${client.user.tag}`);
  const actionPollInterval = Math.max(Number(process.env.ACTION_POLL_INTERVAL_MS || 10000), 5000);
  const initialPollDelay = Math.max(actionPollInterval, 60000);
  const pollSheetActions = async () => {
    await enqueue("シート操作・応募・退職処理", async () => {
      await processSheetActions();
      await processBonusDistribution();
      await processRecruitmentApplications();
      await processTerminations();
    });
    setTimeout(pollSheetActions, actionPollInterval);
  };
  enqueue("起動時全件同期", fullSync)
    .finally(() => setTimeout(pollSheetActions, initialPollDelay));
  console.log(`シート操作・応募・退職処理監視: ${actionPollInterval}ms間隔`);
  console.log(`起動後の初回シート確認: ${initialPollDelay}ms後`);
  console.log(`応募フォーム確認: ${recruitmentPollIntervalMs}ms間隔`);
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
