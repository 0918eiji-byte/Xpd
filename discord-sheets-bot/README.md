# Discord → Googleスプレッドシート同期Bot

## 最短導入（3ステップ）

### 1. GitHubへ置く

この `discord-sheets-bot` フォルダの中身を、そのままGitHubリポジトリへアップロードします。

### 2. RailwayでDeploy

Railwayで **New Project → Deploy from GitHub repo** を選び、このリポジトリを指定します。

### 3. RailwayのVariablesへ5項目を登録

```text
DISCORD_BOT_TOKEN          Discord Botのトークン
DISCORD_GUILD_ID           DiscordサーバーID
ROSTER_ROLE_ID             名簿に表示する人が持つロールID
GOOGLE_SPREADSHEET_ID      1YjZJaEqN5dyqnRZUJ2PNrdkSulE7vOrARVZ_R6iCNcA
GOOGLE_SERVICE_ACCOUNT_JSON  Googleから取得したJSONファイルの中身全部
```

対象ロールを複数にする場合だけ、カンマ区切りで追加できます。

```text
ROSTER_ROLE_IDS=ロールID1,ロールID2
EXCLUDE_ROLE_IDS=休職ロールID,退職ロールID
```

登録後に **Deploy** または **Redeploy** を押します。ログへ次が出れば成功です。

```text
Discord接続完了
全件同期完了
```

## 必須の事前設定

- Discord Developer PortalのBot設定で **SERVER MEMBERS INTENT** をオンにする。
- Googleサービスアカウントのメールアドレスを対象スプレッドシートへ「編集者」で共有する。
- スプレッドシートの「ランク設定」C列へ、各DiscordロールIDを入力する。

## 動作

- `ROSTER_ROLE_ID` を持つメンバーだけ「従業員」へ表示します。
- 複数の階級ロールがある場合、「ランク設定」の優先度が最も小さい階級を採用します。
- ロール変更、加入、脱退をリアルタイム反映します。
- Bot起動時には全メンバーを再確認します。

## 項目やランクの追加

- 「従業員」の列を追加・並べ替えても、Botは見出し名を探すため壊れません。
- Botに必要な見出しがない場合は、2行目の右端へ自動追加します。
- 「ランク設定」へ新しい階級を追加すると、次の同期から自動で査定対象になります。
- ランク設定は最大1000行まで読み込みます。
- 独自の列はBotが上書きしないため、自由に追加して利用できます。

## セキュリティ

`.env`、Botトークン、Google JSON鍵はGitHubへアップロードしないでください。秘密情報はRailwayのVariablesだけへ保存します。
