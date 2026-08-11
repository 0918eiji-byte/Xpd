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

初回導入後は、Discord側のIntent設定が遅れていてもBotが30秒ごとに自動再接続します。GitHub連携のAutodeployを有効にすると、コード更新時のRedeploy操作も不要です。

## 必須の事前設定

- Discord Developer PortalのBot設定で **SERVER MEMBERS INTENT** をオンにする。
- Googleサービスアカウントのメールアドレスを対象スプレッドシートへ「編集者」で共有する。
- スプレッドシートの「ランク設定」C列へ、各DiscordロールIDを入力する。
- 「ランク設定」にランク名 `解雇者` を追加し、C列へ解雇者ロールID、F列へ `はい` を設定する。

## 動作

- `ROSTER_ROLE_ID` を持つメンバーだけ「従業員」へ表示します。
- `ROSTER_ROLE_ID`（Police Officer）だけを持ち、階級ロールがないメンバーは適用ランクを `？？？？` と表示します。
- Discord上でBot以外のロールが1つもないメンバー、またはサーバーから脱退したメンバーは「従業員」から自動削除します。
- 複数の階級ロールがある場合、「ランク設定」の優先度が最も小さい階級を採用します。
- ロール変更、加入、脱退をリアルタイム反映します。
- Bot起動時には全メンバーを再確認します。
- 「従業員」の各行で「変更後ランク」を選び、「実行」のチェックボックスをオンにすると、通常5秒以内にDiscordへ反映します。昇格・降格・解雇を1つの操作欄で扱えます。
- 「適用ランク」を直接変更した場合もDiscordへ反映し、Discord側で階級ロールを変更した場合はスプレッドシートへ反映する双方向同期です。
- 階級ロールが残った状態でPolice Officerロールだけが外れた場合は、BotがPolice Officerロールを自動で付け直します。
- 解雇は名簿対象ロールと階級ロールを解除して「解雇者」ロールを付与し、「解雇者管理」へ氏名・ID・最終ランク・解雇日を自動記録します。
- 「解雇者管理」で「対応署員」を入力して「手続き完了」にチェックすると完了日が自動記録され、その7日後に「従業員」名簿から自動削除されます。退職履歴は「解雇者管理」に残ります。
- 同じDiscordユーザーIDの行が複数ある場合、手入力項目を残しながら起動時に1行へ統合します。

操作の監視間隔を変える場合だけ、Railway Variablesへ次を追加します。未設定時は5000ミリ秒です。

```text
ACTION_POLL_INTERVAL_MS=5000
```

BotのDiscordロールは、操作対象の名簿ロール・階級ロールより上へ配置してください。下にあるロールはDiscordの仕様上変更できません。
「解雇者」ロールもBotのDiscordロールより下へ配置してください。

## 項目やランクの追加

- 「従業員」の列を追加・並べ替えても、Botは見出し名を探すため壊れません。
- Discordの識別には「社員ID」（`DC-DiscordユーザーID`）を使うため、DiscordユーザーID・入社日・手動ランク・同期ランク列は不要です。
- Botに必要な見出しがない場合は、2行目の右端へ自動追加します。
- 「ランク設定」へ新しい階級を追加すると、次の同期から自動で査定対象になります。
- ランク設定は最大1000行まで読み込みます。
- 独自の列はBotが上書きしないため、自由に追加して利用できます。

## セキュリティ

`.env`、Botトークン、Google JSON鍵はGitHubへアップロードしないでください。秘密情報はRailwayのVariablesだけへ保存します。
