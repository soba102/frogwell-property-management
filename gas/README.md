# 不動産管理 AI 自動化モック ─ セットアップ手順

## 稼働中のモック環境

| リソース | URL |
|----------|-----|
| **メインスプレッドシート** | https://docs.google.com/spreadsheets/d/1ysF0GCagzM9-Ch071cmUGgfWzNut8yijpmwRxfco4Rk/edit |
| **Drive（プロジェクトフォルダ）** | https://drive.google.com/drive/folders/1SfsTSLDijkQiZdE7QjYlkjVu58BL_dpe |

コード更新時はこのスプレッドシートに紐づく GAS エディタへ `_PASTE_ME_Code.gs` を貼り付けて保存してください。  
引き継ぎの全体像は [`docs/REAL_ESTATE_AI_ENGINEERING_HANDOFF.md`](../docs/REAL_ESTATE_AI_ENGINEERING_HANDOFF.md) を参照。

---

## 前提条件

- Node.js 20 以上
- Google Workspace アカウント
- Google AI Studio の API キー

## Step 1: clasp のインストールとログイン

```powershell
npm install -g @google/clasp
clasp login
```

ブラウザが開くので、Google アカウントでログインして権限を許可する。

> **注意**: 初めて clasp を使う場合は、Google Apps Script API を有効にする必要があります。
> https://script.google.com/home/usersettings で「Google Apps Script API」を ON にしてください。

## Step 2: GAS プロジェクト作成

```powershell
cd gas
clasp create --type sheets --title "不動産管理_メインデータ"
```

`.clasp.json` が自動生成され、Google Drive にスプレッドシートが作成される。

## Step 3: コードをプッシュ

```powershell
clasp push
```

`appsscript.json` を上書きするか聞かれたら `y` で進む。

## Step 4: スプレッドシートを開く

```powershell
clasp open
```

ブラウザでスプレッドシートが開く。

## Step 5: 初期セットアップ

1. スプレッドシート上部のメニューに「AI管理ツール」が表示される（ページリロードが必要な場合あり）
2. 「AI管理ツール」→「初期セットアップ」を実行
3. 初回実行時に権限の承認ダイアログが出るので許可する

## Step 6: 設定値の入力

「設定」シートに以下を入力する:

| キー名 | 値 |
|--------|------|
| GEMINI_API_KEY | [Google AI Studio](https://aistudio.google.com/apikey) で取得した API キー |
| INVOICE_FOLDER_ID | Google Drive に作成した請求書フォルダの ID（URLの `/folders/` 以降の部分） |
| REPORT_FOLDER_ID | Google Drive に作成したレポート出力フォルダの ID |
| DEPOSIT_FOLDER_ID | Google Drive に作成した入金CSVフォルダの ID（銀行CSVを置く場所） |

### フォルダ ID の確認方法

Google Drive でフォルダを開いたときの URL:
```
https://drive.google.com/drive/folders/1ABCxyz123456789
                                       ^^^^^^^^^^^^^^^^^ ← これがフォルダID
```

## Step 7: テスト実行

1. `sample-invoices/` 内の HTML ファイルをブラウザで開き、PDF として印刷保存する
2. その PDF を Google Drive の請求書フォルダにアップロードする
3. スプレッドシートで「AI管理ツール」→「OCR バッチ実行」を実行
4. 「請求書データ」シートにデータが入っていることを確認
5. 「AI管理ツール」→「レポート生成（全オーナー）」でレポート PDF を生成
6. 「AI管理ツール」→「レポート一括メール送信」でメール送信（テスト時はオーナーマスターのメールアドレスを自分のアドレスに変更しておく）

## Step 8: 入金消込（賃料入金のマッチング）

1. `sample-deposits/sample_bank_statement_2026-06.csv` を Google Drive の入金CSVフォルダ（DEPOSIT_FOLDER_ID）にアップロードする
2. スプレッドシートで「AI管理ツール」→「入金CSV取込（消込）」を実行
3. 「入金消込」シートに、テナントごとの入金状況（消込済 / 一部入金 / 未入金 / 要確認）が色分けで出力される
4. レポートを再生成すると「収入の部」に各テナントの賃料と入金状況が反映される

### 消込の仕組み

- 銀行CSVの**摘要から区画番号（ユニット番号）を抽出**し、「マスター_テナント」の `物件名＋区画番号` と照合してテナント→物件→オーナーを特定（請求書マッチングと同一キー）
  - 例: 摘要 `Condo Setia A-301 RENTAL JUN 2026` → 区画 `A-301` でマッチ
  - 同一区画番号が複数物件にある場合は、摘要内の建物名で絞り込み
  - 区画番号が読めない入金（手数料・JOMPAY等）は家賃以外とみなして無視
- テナントの月間家賃（請求額）と実入金額を突合してステータスを自動判定
  - 一致 → **消込済** / 不足 → **一部入金** / 入金なし → **未入金** / 区画番号が読めたがマスター該当なし → **要確認**

> 「マスター_テナント」は先方の契約情報一覧フォーマット（物件名・区画番号・テナント名・契約期間・更新・解約予告・月間家賃・敷金・備考）に準拠しています。オーナーは「マスター_物件」経由で解決するため、テナント表にオーナー列は不要です。

## マスターのデータモデル（オーナー名キー）

先方の「オーナーリスト」フォーマットに合わせ、**オーナーは名前で識別**します（オーナーID廃止）。

- **マスター_物件**＝先方のオーナーリストそのまま：`物件名 / 部屋番号 / オーナー名 / 住所`
  - 物件は「**物件名（建物名）＋部屋番号**」で特定。請求書OCR・消込・レポートすべて同じキー
  - 同一オーナーが複数物件を持てる（行を分けて記載）
- **マスター_オーナー**＝連絡先のみ：`オーナー名 / メール / 管理手数料率(%) / 備考`
  - 「サンプルマスター投入」時、物件マスターのオーナー名から重複なしで自動生成されます
  - レポート自動送付に使うので **メール列を埋めてください**（空欄のオーナーは送信スキップ）
  - 管理手数料率は空欄ならグローバル設定 `MANAGEMENT_FEE_PCT` を使用

## 請求書の物件マッチング（住所・ユニット番号）

OCR は請求書から **建物名（物件名）・ユニット番号・住所** を抽出し、「**建物名 ＋ ユニット番号**」で
オーナーを特定します（住所だけでは同一建物の別ユニットを区別できないため）。

- ユニット番号は表記ゆれを吸収（`46-5` ＝ `46-05`、`A-36-1` ＝ `a-36-01`、`T1-19-3a` ＝ `T1-19-3A`）
- 建物名がマスターと一致しない場合、検出した建物名を「**マスター_建物別名**」シートへ自動追記します
  - 「正規 建物名」列に正しい建物名を入力 → 次回以降は自動で名寄せされます
- 銀行明細（CIMB等の取引明細）は請求書OCRではなく、消込の元データとして別処理の想定です

## Gemini API キーの取得方法

1. https://aistudio.google.com/apikey にアクセス
2. 「Create API key」をクリック
3. プロジェクトを選択（または新規作成）
4. 表示されるキーをコピーして「設定」シートに貼り付け

## 開発 Tips

### ローカルで編集 → push

```powershell
# コードを編集した後
clasp push

# スプレッドシートのスクリプトエディタを開く
clasp open --addon
```

### ログの確認

```powershell
clasp logs
```

またはスプレッドシートの「拡張機能」→「Apps Script」→「実行数」からログを確認。

## ファイル構成

| ファイル | 役割 |
|----------|------|
| Config.gs | 設定値・定数・ヘルパー関数 |
| Main.gs | メニュー登録・エントリーポイント |
| Setup.gs | 初期セットアップ（シート作成・サンプルデータ） |
| DriveScanner.gs | Drive フォルダの新規ファイル検知 |
| OCR.gs | Gemini API 呼出し・請求書データ構造化 |
| SheetWriter.gs | OCR 結果のスプレッドシート書込み |
| ReportGenerator.gs | オーナー別レポート PDF 生成 |
| EmailSender.gs | レポートメール送信 |
| Reconciliation.gs | 銀行CSV取込・名寄せ・入金消込 |
| Dashboard.gs | ダッシュボード（サマリー）更新 |
