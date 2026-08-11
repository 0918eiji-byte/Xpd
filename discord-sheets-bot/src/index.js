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
  queue = queue.then(work).catch((error) => console.error(`[${label}]`, error));
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
    if (enabled && roleId && rankName) map.set(roleId, { priority, rankName, roleName });
  }
  return map;
}

function assessMember(member, rankMap) {
  const matches = member.roles.cache
    .map((role) => rankMap.get(role.id))
    .filter(Boolean)
    .sort((a, b) => a.priority - b.priority);
  return {
    roleNames: matches.map((item) => item.roleName).join(", "),
    rankName: matches[0]?.rankName || "",
  };
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

const botHeaders = ["社員ID", "表示名", "DiscordユーザーID", "入社日", "雇用状態", "手動ランク", "Discordロール", "同期ランク", "適用ランク", "基本ボーナス", "固定係数", "調整額", "見込ボーナス", "最終同期", "メモ"];

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

async function syncMember(member) {
  if (member.guild.id !== guildId || member.user.bot) return;
  const [rankMap, employeeSheet] = await Promise.all([readRankMap(), readEmployees()]);
  const discordColumn = employeeSheet.headerMap.get("DiscordユーザーID");
  const idColumn = employeeSheet.headerMap.get("社員ID");
  const index = employeeSheet.rows.findIndex((row) => String(row[discordColumn] || "") === member.id);
  const emptyIndex = employeeSheet.rows.findIndex((row) => !row[idColumn]);
  const targetRow = index >= 0 ? index + 3 : emptyIndex >= 0 ? emptyIndex + 3 : employeeSheet.rows.length + 3;
  const hasRosterRole = [...rosterRoleIds].some((id) => member.roles.cache.has(id));
  const hasExcludeRole = [...excludeRoleIds].some((id) => member.roles.cache.has(id));
  const eligible = hasRosterRole && !hasExcludeRole;
  const now = new Date().toISOString();

  if (!eligible) {
    if (index < 0) return;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: rowUpdate(employeeSheet, targetRow, { "雇用状態": hasExcludeRole ? "休職" : "退職", "Discordロール": "", "同期ランク": "", "最終同期": now, "メモ": hasExcludeRole ? "除外ロールあり" : "対象ロールなし" }),
      },
    });
    console.log(`名簿対象外: ${member.displayName}`);
    return;
  }

  const assessed = assessMember(member, rankMap);
  if (index < 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: rowUpdate(employeeSheet, targetRow, { "社員ID": employeeId(member.id), "表示名": member.displayName, "DiscordユーザーID": member.id, "入社日": member.joinedAt?.toISOString() || now, "雇用状態": "在籍", "手動ランク": "", "Discordロール": assessed.roleNames, "同期ランク": assessed.rankName, ...employeeFormulas(employeeSheet, targetRow), "最終同期": now, "メモ": "Railway Bot自動登録" }),
      },
    });
  } else {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: rowUpdate(employeeSheet, targetRow, { "表示名": member.displayName, "DiscordユーザーID": member.id, "雇用状態": "在籍", "Discordロール": assessed.roleNames, "同期ランク": assessed.rankName, "最終同期": now, "メモ": "Railway Botリアルタイム同期" }),
      },
    });
  }
  console.log(`同期完了: ${member.displayName} → ${assessed.rankName || "階級なし"}`);
}

async function fullSync() {
  const guild = await client.guilds.fetch(guildId);
  const members = await guild.members.fetch();
  for (const member of members.values()) await syncMember(member);
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

client.once("ready", () => {
  console.log(`Discord接続完了: ${client.user.tag}`);
  enqueue("起動時全件同期", fullSync);
});
client.on("guildMemberAdd", (member) => enqueue("加入", () => syncMember(member)));
client.on("guildMemberUpdate", (_oldMember, newMember) => enqueue("ロール変更", () => syncMember(newMember)));
client.on("guildMemberRemove", (member) => enqueue("脱退", () => markRemoved(member)));

const port = Number(process.env.PORT || 3000);
http.createServer((_req, res) => {
  res.writeHead(200, { "content-type": "text/plain; charset=utf-8" });
  res.end(client.isReady() ? "Discord Sheets Bot: OK" : "Discord Sheets Bot: starting");
}).listen(port, () => console.log(`Health check: ${port}`));

client.login(process.env.DISCORD_BOT_TOKEN);
