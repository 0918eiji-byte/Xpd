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
    .filter(Boolean)
    .sort((a, b) => a.priority - b.priority);
  return {
    roleNames: [...new Set(matches.map((item) => item.roleName))].join(", "),
    rankName: matches[0]?.rankName || "",
  };
}

function sortedRanks(rankMap) {
  return [...rankMap.values()].sort((a, b) => a.priority - b.priority);
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

const actionHeaders = ["昇格", "降格", "解雇"];
const botHeaders = ["社員ID", "表示名", "DiscordユーザーID", "入社日", "雇用状態", "手動ランク", "Discordロール", "同期ランク", "適用ランク", "基本ボーナス", "固定係数", "調整額", "見込ボーナス", "最終同期", "メモ", ...actionHeaders, "操作結果", "操作日時"];

async function readEmployees() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'従業員'!A2:ZZ1000",
  });
  const values = response.data.values || [];
  let headers = values[0] || [];
  const missingHeaders = botHeaders.filter((header) => !headers.includes(header));
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
  return {
    headers,
    headerMap: new Map(headers.map((header, index) => [String(header), index])),
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
  const status = ref(employeeSheet, "雇用状態", rowNumber);
  const manualRank = ref(employeeSheet, "手動ランク", rowNumber);
  const syncedRank = ref(employeeSheet, "同期ランク", rowNumber);
  const appliedRank = ref(employeeSheet, "適用ランク", rowNumber);
  const base = ref(employeeSheet, "基本ボーナス", rowNumber);
  const factor = ref(employeeSheet, "固定係数", rowNumber);
  const adjustment = ref(employeeSheet, "調整額", rowNumber);
  return {
    "適用ランク": `=IF(${manualRank}<>"",${manualRank},${syncedRank})`,
    "基本ボーナス": `=IF(${id}="","",IFNA(XLOOKUP(${appliedRank},'ランク設定'!$B$3:$B$1000,'ランク設定'!$E$3:$E$1000),0))`,
    "固定係数": `=IF(${id}="","",1)`,
    "調整額": 0,
    "見込ボーナス": `=IF(${id}="","",IF(${status}<>"在籍",0,${base}*${factor}+${adjustment}))`,
  };
}

function employeeId(discordId) {
  return `DC-${discordId}`;
}

function discordIdCell(discordId) {
  return `'${discordId}`;
}

function normalizedDiscordId(row, employeeSheet) {
  const discordIndex = employeeSheet.headerMap.get("DiscordユーザーID");
  const employeeIndex = employeeSheet.headerMap.get("社員ID");
  const direct = String(row[discordIndex] || "").replace(/^'/, "").trim();
  if (direct) return direct;
  return String(row[employeeIndex] || "").replace(/^DC-/, "").trim();
}

function isChecked(value) {
  return value === true || String(value || "").toUpperCase() === "TRUE";
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

async function syncMember(member, context = null) {
  if (member.guild.id !== guildId || member.user.bot) return;
  const [rankMap, employeeSheet] = context
    ? [context.rankMap, context.employeeSheet]
    : await Promise.all([readRankMap(), readEmployees()]);
  const discordColumn = employeeSheet.headerMap.get("DiscordユーザーID");
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const index = employeeSheet.rows.findIndex((row) =>
    String(row[discordColumn] || "") === member.id || String(row[idColumn] || "") === employeeId(member.id),
  );
  const emptyIndex = employeeSheet.rows.findIndex((row) => !row[idColumn]);
  const targetRow = index >= 0 ? index + 3 : emptyIndex >= 0 ? emptyIndex + 3 : employeeSheet.rows.length + 3;
  const hasRosterRole = [...rosterRoleIds].some((id) => member.roles.cache.has(id));
  const hasExcludeRole = [...excludeRoleIds].some((id) => member.roles.cache.has(id));
  const eligible = hasRosterRole && !hasExcludeRole;
  const now = new Date().toISOString();

  if (!eligible) {
    if (index < 0) return;
    await applyEmployeeUpdates(
      rowUpdate(employeeSheet, targetRow, { "雇用状態": hasExcludeRole ? "休職" : "退職", "Discordロール": "", "同期ランク": "", "最終同期": now, "メモ": hasExcludeRole ? "除外ロールあり" : "対象ロールなし" }),
      context,
    );
    console.log(`名簿対象外: ${member.displayName}`);
    return;
  }

  const assessed = assessMember(member, rankMap);
  if (index < 0) {
    await applyEmployeeUpdates(
      rowUpdate(employeeSheet, targetRow, { "社員ID": employeeId(member.id), "表示名": member.displayName, "DiscordユーザーID": discordIdCell(member.id), "入社日": member.joinedAt?.toISOString() || now, "雇用状態": "在籍", "手動ランク": "", "Discordロール": assessed.roleNames, "同期ランク": assessed.rankName, ...employeeFormulas(employeeSheet, targetRow), "最終同期": now, "メモ": "Railway Bot自動登録" }),
      context,
    );
    const rowIndex = targetRow - 3;
    while (employeeSheet.rows.length <= rowIndex) employeeSheet.rows.push([]);
    employeeSheet.rows[rowIndex][idColumn] = employeeId(member.id);
    employeeSheet.rows[rowIndex][discordColumn] = member.id;
  } else {
    await applyEmployeeUpdates(
      rowUpdate(employeeSheet, targetRow, { "表示名": member.displayName, "DiscordユーザーID": discordIdCell(member.id), "雇用状態": "在籍", "Discordロール": assessed.roleNames, "同期ランク": assessed.rankName, "最終同期": now, "メモ": "Railway Botリアルタイム同期" }),
      context,
    );
  }
  console.log(`同期完了: ${member.displayName} → ${assessed.rankName || "階級なし"}`);
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
    "社員ID", "表示名", "DiscordユーザーID", "雇用状態", "Discordロール", "同期ランク",
    "適用ランク", "基本ボーナス", "見込ボーナス", "最終同期", ...actionHeaders, "操作結果", "操作日時",
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
    const noteColumn = employeeSheet.headerMap.get("メモ");
    const currentNote = String(canonicalRow[noteColumn] || "").trim();
    merged["メモ"] = currentNote.includes("重複統合")
      ? currentNote
      : [currentNote, `重複統合: ${indexes.length}件 → 1件`].filter(Boolean).join(" / ");
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
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        sortRange: {
          range: {
            sheetId: 1100459512,
            startRowIndex: 2,
            endRowIndex: 1000,
            startColumnIndex: 0,
            endColumnIndex: Math.max(employeeSheet.headers.length, 21),
          },
          sortSpecs: [{ dimensionIndex: 0, sortOrder: "ASCENDING" }],
        },
      }],
    },
  });
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

async function executeRankAction(guild, member, rankMap, action) {
  const ranks = sortedRanks(rankMap);
  if (!ranks.length) throw new Error("有効なランク設定がありません");
  const currentRanks = ranks.filter((rank) => member.roles.cache.has(rank.roleId));
  const current = currentRanks[0];
  const currentIndex = current ? ranks.findIndex((rank) => rank.roleId === current.roleId) : -1;

  if (action === "解雇") {
    const removeIds = [...new Set([
      ...ranks.filter((rank) => member.roles.cache.has(rank.roleId)).map((rank) => rank.roleId),
      ...[...rosterRoleIds].filter((roleId) => member.roles.cache.has(roleId)),
    ])];
    const blocked = removeIds
      .map((roleId) => guild.roles.cache.get(roleId))
      .filter((role) => role && !role.editable)
      .map((role) => role.name);
    if (blocked.length) throw new Error(`Botより上位のロールは解除できません: ${blocked.join(", ")}`);
    if (removeIds.length) await member.roles.remove(removeIds, "Google Sheetsから解雇");
    return `${current?.rankName || "階級なし"} → 退職`;
  }

  let target;
  if (action === "昇格") {
    if (!current) target = ranks.at(-1);
    else if (currentIndex === 0) throw new Error("すでに最高ランクです");
    else target = ranks[currentIndex - 1];
  } else {
    if (!current) throw new Error("現在のランクロールがありません");
    if (currentIndex === ranks.length - 1) throw new Error("すでに最低ランクです");
    target = ranks[currentIndex + 1];
  }

  const targetRole = guild.roles.cache.get(target.roleId);
  if (!targetRole) throw new Error(`対象ロールが見つかりません: ${target.rankName}`);
  if (!targetRole.editable) throw new Error(`Botより上位のロールは操作できません: ${targetRole.name}`);
  const removeIds = currentRanks
    .filter((rank) => rank.roleId !== target.roleId)
    .map((rank) => rank.roleId);
  const blocked = removeIds
    .map((roleId) => guild.roles.cache.get(roleId))
    .filter((role) => role && !role.editable)
    .map((role) => role.name);
  if (blocked.length) throw new Error(`Botより上位のロールは解除できません: ${blocked.join(", ")}`);
  if (removeIds.length) await member.roles.remove(removeIds, `Google Sheetsから${action}`);
  if (!member.roles.cache.has(target.roleId)) await member.roles.add(target.roleId, `Google Sheetsから${action}`);
  return `${current?.rankName || "階級なし"} → ${target.rankName}`;
}

async function processSheetActions() {
  if (!client.isReady()) return;
  const [guild, rankMap, employeeSheet] = await Promise.all([
    client.guilds.fetch(guildId),
    readRankMap(),
    readEmployees(),
  ]);
  const actionColumns = new Map(actionHeaders.map((header) => [header, employeeSheet.headerMap.get(header)]));

  for (let index = 0; index < employeeSheet.rows.length; index += 1) {
    const row = employeeSheet.rows[index];
    const selected = actionHeaders.filter((header) => isChecked(row[actionColumns.get(header)]));
    if (!selected.length) continue;
    const rowNumber = index + 3;
    const resetFields = Object.fromEntries(actionHeaders.map((header) => [header, false]));

    if (selected.length !== 1) {
      await writeActionResult(employeeSheet, rowNumber, {
        ...resetFields,
        "操作結果": "エラー: 操作は1つだけ選択してください",
        "操作日時": new Date().toISOString(),
      });
      continue;
    }

    const action = selected[0];
    const discordId = normalizedDiscordId(row, employeeSheet);
    if (!discordId) {
      await writeActionResult(employeeSheet, rowNumber, {
        ...resetFields,
        "操作結果": "エラー: DiscordユーザーIDがありません",
        "操作日時": new Date().toISOString(),
      });
      continue;
    }

    await writeActionResult(employeeSheet, rowNumber, { "操作結果": `処理中: ${action}` });
    try {
      const member = await guild.members.fetch(discordId);
      const transition = await executeRankAction(guild, member, rankMap, action);
      await writeActionResult(employeeSheet, rowNumber, {
        ...resetFields,
        "手動ランク": "",
        "雇用状態": action === "解雇" ? "退職" : "在籍",
        "操作結果": `完了: ${action} (${transition})`,
        "操作日時": new Date().toISOString(),
      });
      if (action !== "解雇") await syncMember(await guild.members.fetch(discordId));
      console.log(`シート操作完了: ${member.displayName} ${action} ${transition}`);
    } catch (error) {
      const message = error.response?.data?.message || error.message || String(error);
      await writeActionResult(employeeSheet, rowNumber, {
        ...resetFields,
        "操作結果": `エラー: ${message}`,
        "操作日時": new Date().toISOString(),
      });
      console.error(`シート操作失敗: ${discordId} ${action}`, message);
    }
  }
}

async function fullSync() {
  const guild = await client.guilds.fetch(guildId);
  const members = await guild.members.fetch();
  const missingRosterRoles = [...rosterRoleIds].filter((id) => !guild.roles.cache.has(id));
  if (missingRosterRoles.length) {
    throw new Error(`ROSTER_ROLE_IDがこのサーバーに存在しません: ${missingRosterRoles.join(", ")}`);
  }
  const rosterNames = [...rosterRoleIds].map((id) => guild.roles.cache.get(id)?.name || id);
  const eligibleCount = members.filter((member) =>
    !member.user.bot &&
    [...rosterRoleIds].some((id) => member.roles.cache.has(id)) &&
    ![...excludeRoleIds].some((id) => member.roles.cache.has(id)),
  ).size;
  console.log(`名簿対象ロール: ${rosterNames.join(", ")}`);
  console.log(`名簿対象人数: ${eligibleCount}人`);
  if (eligibleCount === 0) {
    throw new Error("名簿対象者が0人です。ROSTER_ROLE_IDが全署員に共通するロールか確認してください。");
  }
  await consolidateEmployeeDuplicates();
  const [rankMap, employeeSheet] = await Promise.all([readRankMap(), readEmployees()]);
  const context = { rankMap, employeeSheet, pendingData: [] };
  for (const member of members.values()) await syncMember(member, context);
  if (context.pendingData.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data: context.pendingData },
    });
  }
  console.log(`全件同期完了: ${members.size}人確認`);
}

async function markRemoved(member) {
  if (member.guild.id !== guildId || member.user.bot) return;
  const employeeSheet = await readEmployees();
  const discordColumn = employeeSheet.headerMap.get("DiscordユーザーID");
  const index = employeeSheet.rows.findIndex((row) => String(row[discordColumn] || "") === member.id);
  if (index < 0) return;
  const rowNumber = index + 3;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: rowUpdate(employeeSheet, rowNumber, { "雇用状態": "退職", "Discordロール": "", "同期ランク": "", "最終同期": new Date().toISOString(), "メモ": "Discordサーバー脱退" }),
    },
  });
  console.log(`脱退処理: ${member.displayName || member.user.username}`);
}

client.once("clientReady", () => {
  console.log(`Discord接続完了: ${client.user.tag}`);
  enqueue("起動時全件同期", fullSync);
  const actionPollInterval = Math.max(Number(process.env.ACTION_POLL_INTERVAL_MS || 5000), 3000);
  setInterval(() => enqueue("シート操作", processSheetActions), actionPollInterval);
  console.log(`シート操作監視: ${actionPollInterval}ms間隔`);
});
client.on("guildMemberAdd", (member) => enqueue("加入", () => syncMember(member)));
client.on("guildMemberUpdate", (_oldMember, newMember) => enqueue("ロール変更", () => syncMember(newMember)));
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
