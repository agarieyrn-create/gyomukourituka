# GAS 業務効率化ツール集（モノレポ）

Google Apps Script を活用した業務効率化ツールをまとめたモノレポリポジトリです。

## 収録ツール一覧

| ツール名 | ディレクトリ | 概要 |
|----------|-------------|------|
| [請求書・見積書 自動生成ツール](./tools/invoice-generator/) | `tools/invoice-generator/` | スプレッドシートから請求書・見積書を自動生成し、PDF出力する |

## リポジトリ構成

```
gas-tools/
├── README.md                    # このファイル
├── .gitignore
├── tools/
│   ├── invoice-generator/       # 請求書・見積書 自動生成ツール
│   │   ├── src/                 # GASソースコード
│   │   ├── .clasp.json          # clasp設定（git管理外）
│   │   ├── appsscript.json      # GASプロジェクト設定
│   │   └── README.md            # ツール単体の説明
│   └── _template/               # 新ツール追加用テンプレート
│       ├── src/
│       ├── .clasp.json.example
│       ├── appsscript.json
│       └── README.md
└── docs/
    ├── how-to-add-tool.md       # 新ツール追加手順
    └── clasp-setup.md           # clasp初期設定手順
```

## 新しいツールを追加するには

1. `tools/_template/` をコピーして `tools/<ツール名>/` を作成
2. ソースコードを `src/` に配置
3. `.clasp.json` を設定して `clasp push` でデプロイ

詳しくは [docs/how-to-add-tool.md](./docs/how-to-add-tool.md) を参照してください。

## 開発環境

- [clasp](https://github.com/google/clasp) — GASプロジェクトのローカル開発・デプロイ用CLI
- 初期設定は [docs/clasp-setup.md](./docs/clasp-setup.md) を参照
