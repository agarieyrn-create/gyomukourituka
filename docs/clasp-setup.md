# clasp 初期設定手順（開発者向け）

[clasp](https://github.com/google/clasp) は Google Apps Script をローカルで開発・管理するための CLI ツールです。

## 前提条件

- Node.js（v12以上）がインストール済みであること

## インストール

```bash
npm install -g @google/clasp
```

## Googleアカウントでログイン

```bash
clasp login
```

ブラウザが開くので、Googleアカウントでログインして権限を許可します。

## GAS APIを有効化

1. [Google Apps Script API 設定ページ](https://script.google.com/home/usersettings) を開く
2. 「Google Apps Script API」を **オン** にする

## プロジェクトとの紐付け

### 既存のGASプロジェクトに接続する場合

1. 対象ツールのディレクトリに移動

```bash
cd tools/<ツール名>
```

2. `.clasp.json` を作成（テンプレートからコピー）

```bash
cp .clasp.json.example .clasp.json
```

3. `scriptId` を書き換え

```json
{
  "scriptId": "対象のスクリプトID",
  "rootDir": "src"
}
```

### 新規GASプロジェクトを作成する場合

```bash
cd tools/<ツール名>
clasp create --type sheets --rootDir src
```

作成後、生成された `.clasp.json` を確認してください。

## よく使うコマンド

| コマンド | 説明 |
|---------|------|
| `clasp push` | ローカルのコードをGASプロジェクトにアップロード |
| `clasp pull` | GASプロジェクトのコードをローカルにダウンロード |
| `clasp open` | GASエディタをブラウザで開く |
| `clasp status` | push対象のファイルを確認 |
| `clasp logs` | 実行ログを表示 |

## 注意事項

- `.clasp.json` には `scriptId` が含まれるため、`.gitignore` で管理対象外にしています
- 各ツールごとに個別の `.clasp.json` を持ち、それぞれ異なるGASプロジェクトに紐付けます
- `clasp push` は `rootDir` で指定したディレクトリ（通常は `src/`）配下のファイルをアップロードします
