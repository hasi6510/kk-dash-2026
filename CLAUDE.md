# 家計ダッシュボード — Claude への指示

このファイルはClaude Code が自動で読み込む指示書です。

---

## ライフプランシートへの月次転記

たっくーさんから「先月のライフプランシートを作って」と依頼されたら**必ず以下を実行すること**：

### 手順

1. MoneyForwardで先月データを取得
   - Chrome MCPで `https://moneyforward.com/cf/summary?month=YYYY-MM` を開く（YYY-MMは先月）
   - 収支内訳の各カテゴリ金額を読み取る

2. 以下のカテゴリマッピングで変換する

| MoneyForward | ライフプランシート（詳細列） |
|-------------|--------------------------|
| 当月収入（合計） | 給料 |
| 住宅 / 家賃 | 家賃 |
| 住宅 / 通信費 | 通信費 |
| 住宅 / 水道代光熱費 | 水道光熱費 |
| 住宅 / サブスク | サブスク(毎月) |
| 未分類（リベシティ相当） | リベシティ |
| 食費カテゴリ / 食費 | 食費 |
| 食費カテゴリ / 趣味・自己投資 | 趣味・自己投資 |
| 食費カテゴリ / 日用品費 | 日用品 |
| 食費カテゴリ / 美容費 | 美容院・脱毛サロン |
| 食費カテゴリ / 交通費 | 交通費 |
| 趣味娯楽 / 病気治療費 | 医療費 |

3. Google Drive に `mf_monthly_data.json` を **create_file** で作成する（毎月新規作成）

```json
{
  "month": "YYYY-MM",
  "created_at": "<ISO timestamp>",
  "entries": [
    { "label": "給料",               "value": 000000 },
    { "label": "家賃",               "value": 000000 },
    { "label": "サブスク(毎月)",     "value": 000000 },
    { "label": "通信費",             "value": 000000 },
    { "label": "水道光熱費",         "value": 000000 },
    { "label": "リベシティ",         "value": 000000 },
    { "label": "食費",               "value": 000000 },
    { "label": "趣味・自己投資",     "value": 000000 },
    { "label": "日用品",             "value": 000000 },
    { "label": "美容院・脱毛サロン", "value": 000000 },
    { "label": "交通費",             "value": 000000 },
    { "label": "医療費",             "value": 000000 }
  ]
}
```

4. たっくーさんへ伝えること
   > `mf_monthly_data.json` を作成しました！  
   > ライフプランシートの Apps Script で「▶ 実行」を押すと自動で転記されます📊

### 関連情報
- ライフプランシートID: `1vztbd3xPg-Y5XheaLOlxdzbs8plIMFgoAtUrmlI1GUc`
- Apps Script の関数名: `enterMonthlyData`
- Google Drive JSONファイル名: `mf_monthly_data.json`

---

## 翌月の予算を計算・確定したとき

たっくーさんと来月の予算を相談して金額が決まったら、**必ず以下を実行すること**：

### 1. `budget.json` を更新する

```json
{
  "month": "YYYY-MM",  ← 翌月に変更（例: "2026-06"）
  "variable_categories": [
    { "name": "食費",           "budget": XX000 },
    { "name": "日用品",         "budget": XX000 },
    { "name": "美容",           "budget": XX000 },
    { "name": "交通費",         "budget": XX000 },
    { "name": "趣味・自己投資", "budget": XX000 },
    { "name": "医療費",         "budget": XX000 },
    { "name": "外食・カフェ",   "budget": XX000 },
    { "name": "その他",         "budget": XX000 },
    { "name": "転職活動費",     "budget": XX000 },
    { "name": "婚活費用",       "budget": XX000 }
  ],
  "income_monthly": XXXXXX  ← 収入が変わった場合は更新
}
```

### 2. 更新後にたっくーさんへ伝えること

> `budget.json` を [翌月] 用に更新しました！  
> `run.bat` をダブルクリックすると新しい予算でダッシュボードが表示されます 📊

---

## サブスクを追加・解約したとき

Googleスプレッドシートのサブスクシートを更新してから：

1. 新しいタブを追加（例: "2026年改善ver3"）して内容を更新
2. ファイル → ダウンロード → Microsoft Excel (.xlsx)
3. ファイル名を `subscriptions.xlsx` に変更
4. `C:\Users\ausbr\Desktop\Claude\Budget-Dashboard\` に保存（上書きOK）
5. `run.bat` を実行 → 自動で最新 verX タブを検出し、**終了日が過去のサービスを自動除外**

**タブ命名規則:** "2026年改善ver2", "2026年改善ver3" など `verX` パターン → 数字が最大のタブを使用

---

## 高配当株を買い増ししたとき

1. Googleスプレッドシートを最新に更新
2. ファイル → ダウンロード → Microsoft Excel (.xlsx) でダウンロード
3. ファイル名を `stocks.xlsx` に変更
4. `C:\Users\ausbr\Desktop\Claude\Budget-Dashboard\` に保存（上書きOK）
5. `run.bat` を実行 → 自動で最新シートから配当・元本データを読み込む

**シート名規則:** 日付が含まれていれば自動検出（例: "2026-05", "2026年5月"）

---

## 投資元本を更新するとき

`budget.json` の `investment_principal` を実際の取得金額合計に更新する。

```json
"investment_principal": 3800000  ← 実際の累計投資元本（円）に変更
```

これを正確にすると r の計算（年間配当 ÷ 投資元本）が正確になる。

---

## サブスク内訳を更新するとき

`budget.json` の `subscriptions_detail` を編集：

```json
"subscriptions_detail": [
  { "name": "サービス名", "amount": 月額 },
  ...
]
```

合計が `fixed_costs` の `サブスク` 金額と一致するようにすること。

---

## 家計ダッシュボードのCSV更新手順

毎月 MoneyForward ME からCSVをダウンロードしてダッシュボードを更新するとき：

### 手順
1. [MoneyForward ME](https://moneyforward.com/cf) にログイン
2. **家計簿 → 収支内訳 → CSVダウンロード**（対象月を選択）
3. ダウンロードしたCSVファイルを `mf_csv/` フォルダに保存
4. `run.bat` をダブルクリック → ダッシュボード更新完了！

### 自動更新（Task Scheduler）
`setup_scheduler.ps1` を管理者として実行すると **毎月10日 08:00** に自動で `run_silent.bat` が実行されます。
CSVを `mf_csv/` フォルダに入れておけば、10日朝に自動更新されます。

### mf_csv フォルダについて
- 複数のCSVがあっても最新のファイルを自動で使用します
- ファイル名はMoneyForwardのデフォルトのままでOKです（例: `収支明細_202604.csv`）

---

## ファイル構成

```
Budget-Dashboard/
├── index.html          ← ブラウザで開くダッシュボード
├── budget.json         ← 毎月更新する予算設定（Claude が更新）
├── data.json           ← 自動生成（手動編集不要）
├── data_inline.js      ← 自動生成（手動編集不要）
├── update.js           ← Node.jsデータ処理スクリプト
├── run.bat             ← ダブルクリックで更新
├── stocks.xlsx         ← 高配当株シート（買い増し時に置き換え）
├── setup_scheduler.ps1 ← タスクスケジューラ設定（初回のみ）
├── mf_csv/             ← MoneyForward CSVをここに置く
└── CLAUDE.md           ← この指示書
```

---

## カテゴリ追加のルール

新しいカテゴリを追加する場合：
1. `variable_categories` に `{ "name": "カテゴリ名", "budget": 金額 }` を追加
2. `mf_category_map` にMoneyForwardのカテゴリ名とのマッピングを追加
3. `node update.js` を実行して動作確認
