# 不動産管理 AI 自動化

マレーシアの不動産管理業務（請求書 OCR・オーナー別レポート・入金消込）を **Google Apps Script + Gemini** で自動化するプロトタイプ／本実装用リポジトリです。

## モック環境（稼働中）

| リソース | URL |
|----------|-----|
| **メインスプレッドシート** | https://docs.google.com/spreadsheets/d/1ysF0GCagzM9-Ch071cmUGgfWzNut8yijpmwRxfco4Rk/edit |
| **Google Drive** | https://drive.google.com/drive/folders/1SfsTSLDijkQiZdE7QjYlkjVu58BL_dpe |

## エンジニア向け：まず読むもの

1. **[エンジニア引き継ぎドキュメント](./docs/REAL_ESTATE_AI_ENGINEERING_HANDOFF.md)** — 全体像・実装状況・本番ロードマップ
2. **[GAS セットアップ手順](./gas/README.md)** — コード反映・動作確認
3. **[ヒアリングリスト](./docs/06_第一弾_ヒアリングリスト.md)** — クライアント確認事項（回答用 CSV 同梱）

## リポジトリ構成

```
├── gas/                    GAS ソース（モジュール分割 + _PASTE_ME_Code.gs 統合版）
├── docs/                   提案資料・ヒアリング・引き継ぎ
├── sample-invoices/        テスト用請求書 HTML
└── sample-deposits/        テスト用銀行 CSV
```

## クイックスタート（コード更新）

1. 上記スプレッドシートの **拡張機能 → Apps Script** を開く
2. `gas/_PASTE_ME_Code.gs` の全文を貼り付けて保存
3. スプレッドシートをリロード → メニュー **「AI管理ツール」** から実行

## 主な機能（第一弾）

- 請求書 PDF/画像の Gemini OCR → スプレッドシート転記
- 物件名＋部屋番号によるオーナー紐付け（別名マスター・ユニット一意フォールバック）
- 銀行 CSV からの家賃入金消込（区画番号ベース）
- オーナー別月次レポート PDF 生成・Gmail 一括送信

## 関連資料

| ファイル | 内容 |
|----------|------|
| `docs/01_業務フロー確認書.html` | 現状業務フロー（クライアント確認用） |
| `docs/04_プロトタイプ機能一覧.html` | モック機能説明 |
| `docs/05_問い合わせ自動化_企画.html` | Phase 2 企画（未実装） |

---

*旧 `frogwell_sandbox` リポジトリから不動産案件のみ分離しました。*
