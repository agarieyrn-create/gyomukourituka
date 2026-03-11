# 新しいGASツールの追加手順

## 手順

### 1. テンプレートをコピー

```bash
cp -r tools/_template tools/<ツール名>
```

例: 勤怠管理ツールを追加する場合

```bash
cp -r tools/_template tools/attendance-tracker
```

### 2. ソースコードを作成

`tools/<ツール名>/src/` 配下にGASのソースコードを配置します。

- `Code.gs` — メイン処理（テンプレートの雛形を編集）
- 必要に応じて追加の `.gs` ファイルを作成

### 3. appsscript.json を編集

ツールが必要とするOAuthスコープを追加します。

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
  ]
}
```

よく使うスコープ:

| スコープ | 用途 |
|---------|------|
| `https://www.googleapis.com/auth/spreadsheets` | スプレッドシートの読み書き |
| `https://www.googleapis.com/auth/drive` | Googleドライブへのアクセス |
| `https://www.googleapis.com/auth/gmail.send` | メール送信 |
| `https://www.googleapis.com/auth/calendar` | カレンダー操作 |
| `https://www.googleapis.com/auth/script.external_request` | 外部APIへのリクエスト |

### 4. clasp を設定

```bash
cd tools/<ツール名>
cp .clasp.json.example .clasp.json
```

`.clasp.json` の `scriptId` を対象のGASプロジェクトIDに書き換えます。

```json
{
  "scriptId": "実際のスクリプトID",
  "rootDir": "src"
}
```

スクリプトIDの確認方法:
1. GASエディタを開く（スプレッドシート → 拡張機能 → Apps Script）
2. 左サイドバーの「プロジェクトの設定」を選択
3. 「スクリプト ID」に表示されている文字列をコピー

### 5. コードをデプロイ

```bash
cd tools/<ツール名>
clasp push
```

### 6. ルートの README.md を更新

`README.md` の「収録ツール一覧」テーブルに新しいツールを追加します。

### 7. コミット & プッシュ

```bash
git add tools/<ツール名>
git commit -m "Add <ツール名>"
git push
```
