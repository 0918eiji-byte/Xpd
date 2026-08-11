import http from "node:http";
import { Client, GatewayIntentBits } from "discord.js";
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
  intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMembers],
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
const botHeaders = ["社員ID", "表示名", "Discordロール", "適用ランク", "基本ボーナス", "固定係数", "調整額", "見込ボーナス", rankSelectionHeader, actionTriggerHeader, "操作結果", "操作日時", sortPriorityHeader];
const terminationHeaders = ["社員ID", "表示名", "DiscordユーザーID", "最終ランク", "解雇日", "手続き完了", "対応署員", "完了日", "名簿削除予定日", "名簿削除状況", "備考"];
const terminationSheetName = "解雇者管理";
const employeeSheetId = 1100459512;
const retentionPeriodMs = 7 * 24 * 60 * 60 * 1000;

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
  const discordRoles = ref(employeeSheet, "Discordロール", rowNumber);
  const appliedRank = ref(employeeSheet, "適用ランク", rowNumber);
  const base = ref(employeeSheet, "基本ボーナス", rowNumber);
  const factor = ref(employeeSheet, "固定係数", rowNumber);
  const adjustment = ref(employeeSheet, "調整額", rowNumber);
  return {
    "基本ボーナス": `=IF(${id}="","",IFNA(XLOOKUP(${appliedRank},'ランク設定'!$B$3:$B$1000,'ランク設定'!$E$3:$E$1000),0))`,
    "固定係数": `=IF(${id}="","",1)`,
    "調整額": 0,
    "見込ボーナス": `=IF(${id}="","",IF(${discordRoles}="解雇者",0,${base}*${factor}+${adjustment}))`,
  };
}

function employeeId(discordId) {
  return `DC-${discordId}`;
}

async function readTerminations() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${terminationSheetName}'!A2:K500`,
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
    fields["解雇日"] = new Date().toISOString();
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
    "基本ボーナス", "見込ボーナス", rankSelectionHeader, actionTriggerHeader, "操作結果", "操作日時", sortPriorityHeader,
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

async function processTerminations() {
  const [terminationSheet, employeeSheet] = await Promise.all([readTerminations(), readEmployees()]);
  const completeColumn = terminationSheet.headerMap.get("手続き完了");
  const handlerColumn = terminationSheet.headerMap.get("対応署員");
  const completedAtColumn = terminationSheet.headerMap.get("完了日");
  const deletionStatusColumn = terminationSheet.headerMap.get("名簿削除状況");
  const noteColumn = terminationSheet.headerMap.get("備考");
  const terminationEmployeeIdColumn = terminationSheet.headerMap.get("社員ID");
  const terminationDiscordIdColumn = terminationSheet.headerMap.get("DiscordユーザーID");
  const employeeIdColumn = employeeSheet.headerMap.get("社員ID");
  const now = new Date();
  const terminationUpdates = [];
  const employeeRangesToClear = [];

  for (let index = 0; index < terminationSheet.rows.length; index += 1) {
    const row = terminationSheet.rows[index];
    if (!String(row[terminationEmployeeIdColumn] || "").trim()) continue;
    if (!isChecked(row[completeColumn])) continue;
    if (!String(row[handlerColumn] || "").trim()) continue;
    if (String(row[deletionStatusColumn] || "").trim() === "削除済") continue;

    const rowNumber = index + 3;
    const completedAtText = String(row[completedAtColumn] || "").trim();
    if (!completedAtText) {
      terminationUpdates.push(...terminationRowUpdate(terminationSheet, rowNumber, { "完了日": now.toISOString() }));
      console.log(`退職手続き完了日を記録: ${row[terminationEmployeeIdColumn]}`);
      continue;
    }

    const completedAt = new Date(completedAtText);
    if (Number.isNaN(completedAt.getTime()) || now.getTime() < completedAt.getTime() + retentionPeriodMs) continue;

    const discordId = String(row[terminationDiscordIdColumn] || "").replace(/^'/, "").trim();
    const terminatedEmployeeId = String(row[terminationEmployeeIdColumn] || "").trim();
    const employeeIndex = employeeSheet.rows.findIndex((employeeRow) => {
      const candidateDiscordId = normalizedDiscordId(employeeRow, employeeSheet);
      return (discordId && candidateDiscordId === discordId)
        || String(employeeRow[employeeIdColumn] || "").trim() === terminatedEmployeeId;
    });
    if (employeeIndex >= 0) employeeRangesToClear.push(`'従業員'!A${employeeIndex + 3}:ZZ${employeeIndex + 3}`);
    const currentNote = String(row[noteColumn] || "").trim();
    const deletionNote = `退職手続き完了から7日経過のため名簿から自動削除: ${now.toISOString()}`;
    terminationUpdates.push(...terminationRowUpdate(terminationSheet, rowNumber, {
      "名簿削除状況": "削除済",
      "備考": [currentNote, deletionNote].filter(Boolean).join(" / "),
    }));
    console.log(`名簿自動削除: ${terminatedEmployeeId}`);
  }

  if (employeeRangesToClear.length) {
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId,
      requestBody: { ranges: [...new Set(employeeRangesToClear)] },
    });
    await sortEmployees(employeeSheet);
  }
  if (terminationUpdates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data: terminationUpdates },
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
  enqueue("起動時全件同期", fullSync);
  const actionPollInterval = Math.max(Number(process.env.ACTION_POLL_INTERVAL_MS || 10000), 5000);
  setInterval(() => enqueue("シート操作・退職処理", async () => {
    await processSheetActions();
    await processTerminations();
  }), actionPollInterval);
  console.log(`シート操作・退職処理監視: ${actionPollInterval}ms間隔`);
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
