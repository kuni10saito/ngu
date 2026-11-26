---
name: receipt-to-csv
description: Use this skill when the user asks to "read a receipt", "scan a receipt", "extract receipt data", "レシートを読み込む", "レシートをCSVに書き込む", "レシート画像をCSV", or wants to process receipt images into CSV format. Especially useful for オフィス&ホームIT・AI部門 の仕入れレシート（パソコン・部品・消耗品の購入）や財務AIの経費記録に。
disable-model-invocation: false
allowed-tools: Read, Write, Bash, Glob
---

# レシート画像 → CSV 書き込みスキル

ユーザーがレシート画像を指定した場合、画像を読み込んでデータを抽出し、CSVファイルに追記・作成する。
パコラボ社内では主にオフィス&ホームIT・AI部門の仕入れ（PC本体・部品・消耗品）や、各部門の経費精算の記録に利用する。

## 実行手順

### 1. 引数の解釈

`$ARGUMENTS` を確認する：
- 画像パスが含まれている場合 → そのパスを使用
- CSVパスが含まれている場合 → そのパスに書き込む
- 引数なし → ユーザーに画像パスを確認する

パターン例：
- `/receipt-to-csv` → ユーザーにパスを尋ねる
- `/receipt-to-csv ~/receipts/photo.jpg` → その画像を処理
- `/receipt-to-csv ~/receipts/photo.jpg ~/pacolabo/finance/expenses.csv` → 指定CSVに書き込む

### 2. 画像を読み込む

Read ツールで画像ファイルを開く（PNG / JPG / HEIC などすべて対応）。

### 3. レシートデータを抽出する

画像から以下の項目を読み取る：

| フィールド | 説明 | 例 |
|---|---|---|
| date | 購入日 (YYYY-MM-DD) | 2026-05-14 |
| store_name | 店舗名 | 〇〇電器店 扶桑店 |
| item_name | 商品名 | USBメモリ 32GB |
| quantity | 数量 | 1 |
| unit_price | 単価（税込） | 1,200 |
| total_price | 合計金額 | 1,200 |
| tax_8 | 8%対象税額 | 0 |
| tax_10 | 10%対象税額 | 109 |
| payment_method | 支払方法 | 現金 / クレジット / 電子マネー |
| note | 備考（部門・用途など） | オフィス&ホームIT部門 仕入れ |

- 読み取れない項目は空欄にする
- 商品が複数ある場合は **1商品1行** で出力する

### 4. CSVファイルに書き込む

#### CSVパスの決定
- ユーザー指定がある場合 → そのパスを使用
- 指定なし → `~/pacolabo/receipts/expenses.csv` をデフォルトとして使用

#### ファイルが存在しない場合
以下のヘッダー行を含む新規CSVを作成する：

```
date,store_name,item_name,quantity,unit_price,total_price,tax_8,tax_10,payment_method,note
```

#### ファイルが存在する場合
1. Read ツールで既存CSVを確認する
2. ヘッダーが一致することを確認する
3. データ行を **末尾に追記** する（ヘッダーは書かない）

#### CSV出力形式
- 区切り文字: カンマ `,`
- 文字コード: UTF-8
- 日付: `YYYY-MM-DD`
- 金額: 数字のみ（円記号・カンマ不要）
- スペースを含む文字列はダブルクォートで囲む
- 不明な値は空欄（フィールドは省略しない）

#### 出力例
```csv
date,store_name,item_name,quantity,unit_price,total_price,tax_8,tax_10,payment_method,note
2026-05-14,〇〇電器店 扶桑店,USBメモリ 32GB,1,1200,1200,0,109,クレジット,オフィス&ホームIT部門 仕入れ
2026-05-14,〇〇電器店 扶桑店,LANケーブル 5m,2,500,1000,0,91,クレジット,オフィス&ホームIT部門 仕入れ
```

### 5. 完了報告

以下をユーザーに報告する：
- 書き込んだCSVのパス
- 追記した行数
- 抽出したデータのサマリー（店舗・日付・合計金額）
- 読み取れなかった項目があれば明示する
- 経費記録として財務AIに引き渡す場合は、その旨を併せて伝える

## エラー処理

- 画像が存在しない → パスを確認するようユーザーに伝える
- レシートが不鮮明 → 読み取れた項目のみ書き込み、不明項目を報告する
- 画像がレシートでない → その旨を伝えて処理を中止する
