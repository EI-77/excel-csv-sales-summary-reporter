# Excel CSV Sales Summary Reporter

複数の売上CSV・Excelファイルを読み込み，商品マスタと照合して，月別・カテゴリ別・商品別の売上集計レポートを出力するPythonツールです．

`input` フォルダに配置した複数の売上データを読み込み，`master` フォルダに配置した商品マスタと照合します．
商品コードがマスタに存在する行は，商品名，カテゴリなどを補完し，売上金額を計算したうえで `sales_summary_report.xlsx` に出力します．
商品コードがマスタに存在しない行や，売上データ自体に不備がある行は，専用シートに分離して出力します．

このツールは，単なるExcel集計ツールではなく，Excel業務でよく行われる売上データ集計，商品マスタ照合，エラー行抽出，レポート作成を自動化するためのツールです．

---

## 主な機能

* 売上CSVファイルの読み込み
* 売上Excelファイルの読み込み
* 複数売上ファイルの一括読み込み
* 商品マスタCSVファイルの読み込み
* 商品マスタExcelファイルの読み込み
* 売上データと商品マスタの照合
* 商品名・カテゴリの補完
* 数量と単価による売上金額の計算
* 月別売上集計
* カテゴリ別売上集計
* 商品別売上集計
* 売上明細のExcel出力
* 商品マスタに存在しない商品コードの抽出
* 売上データの必須項目チェック
* 売上データの数値チェック
* 売上データの日付形式変換
* 商品マスタの必須項目チェック
* 商品マスタの商品コード重複チェック
* 商品マスタの商品コード空欄チェック
* 処理結果サマリーのExcel出力
* 複数シート構成のExcelレポート出力
* 初回実行用サンプルファイルの自動作成
* 日本語CSV向けの文字コード対応

---

## 使用技術

* Python
* pandas
* openpyxl

---

## 想定用途

このツールは，以下のような作業を自動化するためのものです．

* 複数月の売上データをまとめて集計する
* CSV・Excelの売上明細を一括で処理する
* 商品コードから商品名やカテゴリを自動補完する
* 売上データと商品マスタを照合する
* 商品マスタに存在しない商品コードを抽出する
* 売上データの入力ミスを確認する
* 月別の売上合計を確認する
* カテゴリ別の売上合計を確認する
* 商品別の売上合計を確認する
* ExcelのピボットテーブルやVLOOKUP作業を自動化する
* 売上レポート作成業務を効率化する
* CSV・Excel業務自動化案件の補助ツールとして使う

---

## ディレクトリ構成

```text
excel-csv-sales-summary-reporter/
├─ excel_csv_sales_summary_reporter.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  ├─ sample_sales_2026_06.xlsx
│  └─ sample_sales_2026_07.xlsx
├─ master/
│  └─ product_master.xlsx
└─ output/
   └─ sales_summary_report.xlsx
```

### フォルダの役割

| フォルダ     | 内容                            |
| -------- | ----------------------------- |
| `input`  | 集計対象の売上CSV・Excelファイルを配置するフォルダ |
| `master` | 商品マスタCSV・Excelファイルを配置するフォルダ   |
| `output` | 売上集計レポートの出力先                  |

---

## 実行方法

### 1. リポジトリを取得

```bash
git clone <repository_url>
```

### 2. ディレクトリへ移動

```bash
cd excel-csv-sales-summary-reporter
```

### 3. 必要ライブラリをインストール

```bash
pip install -r requirements.txt
```

### 4. 売上データを配置

集計したい売上データを `input` フォルダに配置します．

対応形式は以下です．

```text
.csv
.xlsx
```

例：

```text
input/sample_sales_2026_06.xlsx
input/sample_sales_2026_07.xlsx
```

CSVファイルでも使用できます．

```text
input/sample_sales_2026_06.csv
input/sample_sales_2026_07.csv
```

### 5. 商品マスタを配置

商品マスタを `master` フォルダに配置します．

例：

```text
master/product_master.xlsx
```

CSVファイルでも使用できます．

```text
master/product_master.csv
```

### 6. config.json を確認

`config.json` で，売上ファイル名，商品マスタファイル名，照合キー，必須列，数値列，日付列，売上金額計算などを設定します．

### 7. プログラムを実行

```bash
python excel_csv_sales_summary_reporter.py
```

または

```bash
py excel_csv_sales_summary_reporter.py
```

### 8. 出力ファイルを確認

実行後，`output` フォルダに以下のExcelレポートが出力されます．

```text
output/
└─ sales_summary_report.xlsx
```

---

## 初回実行について

初回実行時，以下のファイルやフォルダが存在しない場合は自動作成されます．

```text
config.json
input/sample_sales_2026_06.xlsx
input/sample_sales_2026_07.xlsx
master/product_master.xlsx
output/
```

そのため，最初は自分で入力ファイルを用意しなくても，以下のコマンドだけで動作確認できます．

```bash
python excel_csv_sales_summary_reporter.py
```

サンプルデータには，正常に集計できる行，商品マスタに存在しない行，商品コードが空欄の行，数量が数値ではない行，単価が数値ではない行，日付が不正な行が含まれています．
これにより，売上明細，集計結果，入力値エラー，マスタ未登録データが分かれて出力される動作を確認できます．

---

## 入力データの形式

## 売上データ

売上データは，CSVまたはExcelの表形式を想定しています．

* 1行目に列名があること
* 2行目以降に売上データがあること
* Excelファイルの場合，最初のシートを読み込むこと
* `config.json` の `sales_required_columns` に指定した列が存在すること
* `sales_key` に指定した列が存在すること
* `sales_numeric_columns` に指定した列が数値として扱えること
* `sales_date_columns` に指定した列が日付として扱えること

サンプルの売上データ：

| sales_id     | sales_date | product_code | quantity | unit_price | note           |
| ------------ | ---------- | ------------ | -------: | ---------: | -------------- |
| S-202606-001 | 2026/06/01 | P-1001       |        2 |       1800 | 正常データ          |
| S-202606-002 | 2026/06/02 | P-1002       |        1 |       3200 | 正常データ          |
| S-202606-003 | 2026/06/03 | P-1004       |        1 |       9800 | 正常データ          |
| S-202606-004 | 2026/06/04 | P-9999       |        3 |       1000 | マスタに存在しない商品コード |
| S-202606-005 | 2026/06/05 | P-1003       |      abc |       4500 | 数量が数値ではない      |
| S-202606-006 | not-date   | P-1001       |        1 |       1800 | 日付が不正          |

`note` はサンプル用の備考列です．
処理には必須ではありません．

---

## 商品マスタ

商品マスタは，CSVまたはExcelの表形式を想定しています．

* 1行目に列名があること
* 2行目以降に商品マスタがあること
* Excelファイルの場合，最初のシートを読み込むこと
* `config.json` の `master_required_columns` に指定した列が存在すること
* `master_key` に指定した列が存在すること
* `master_key` の値が空欄でないこと
* `master_key` の値が重複していないこと

サンプルの商品マスタ：

| product_code | product_name | category |
| ------------ | ------------ | -------- |
| P-1001       | ワイヤレスマウス     | PC周辺機器   |
| P-1002       | USBキーボード     | PC周辺機器   |
| P-1003       | Webカメラ       | PC周辺機器   |
| P-1004       | 外付けSSD       | ストレージ    |

商品マスタは，売上データに商品名やカテゴリを補完するための基準データです．
そのため，商品マスタに不備がある場合は，売上行単位のエラーとして扱わず，処理全体を停止します．

---

## config.json

動作設定は `config.json` で変更できます．

### サンプル設定

```json
{
  "sales_files": [
    "sample_sales_2026_06.xlsx",
    "sample_sales_2026_07.xlsx"
  ],
  "master_file": "product_master.xlsx",
  "report_file": "sales_summary_report.xlsx",
  "sales_key": "product_code",
  "master_key": "product_code",
  "sales_required_columns": [
    "sales_id",
    "sales_date",
    "product_code",
    "quantity",
    "unit_price"
  ],
  "master_required_columns": [
    "product_code",
    "product_name",
    "category"
  ],
  "master_value_columns": [
    "product_name",
    "category"
  ],
  "sales_numeric_columns": [
    "quantity",
    "unit_price"
  ],
  "sales_date_columns": [
    "sales_date"
  ],
  "date_format": "%Y-%m-%d",
  "month_format": "%Y-%m",
  "amount_calculation": {
    "enabled": true,
    "quantity_column": "quantity",
    "price_column": "unit_price",
    "output_column": "sales_amount"
  }
}
```

---

## 設定項目

| 項目                        | 内容                             |
| ------------------------- | ------------------------------ |
| `sales_files`             | `input` フォルダ内で読み込む売上データファイル一覧  |
| `master_file`             | `master` フォルダ内で読み込む商品マスタファイル   |
| `report_file`             | `output` フォルダに出力する売上集計レポートファイル |
| `sales_key`               | 売上データ側の照合キー列                   |
| `master_key`              | 商品マスタ側の照合キー列                   |
| `sales_required_columns`  | 売上データで空欄を許可しない列                |
| `master_required_columns` | 商品マスタで空欄を許可しない列                |
| `master_value_columns`    | 商品マスタから補完する列                   |
| `sales_numeric_columns`   | 売上データで数値として扱う列                 |
| `sales_date_columns`      | 売上データで日付として扱う列                 |
| `date_format`             | 日付の出力形式                        |
| `month_format`            | 月別集計で使用する年月の出力形式               |
| `amount_calculation`      | 売上金額計算の設定                      |

---

## 照合キーについて

このツールでは，`sales_key` と `master_key` に指定した列を使って，売上データと商品マスタを照合します．

標準設定では，どちらも `product_code` です．

```json
{
  "sales_key": "product_code",
  "master_key": "product_code"
}
```

この場合，以下のように照合します．

```text
売上データ.product_code
        ↓
商品マスタ.product_code
```

商品コードが一致した場合，商品マスタから `product_name`，`category` などの情報を補完します．

---

## master_value_columns について

`master_value_columns` は，商品マスタから売上データへ補完する列を指定する設定です．

例：

```json
{
  "master_value_columns": [
    "product_name",
    "category"
  ]
}
```

この場合，商品コードが一致した売上行に対して，商品マスタの以下の情報が追加されます．

```text
product_name
category
```

---

## amount_calculation について

`amount_calculation` は，数量と単価から売上金額を計算する設定です．

例：

```json
{
  "amount_calculation": {
    "enabled": true,
    "quantity_column": "quantity",
    "price_column": "unit_price",
    "output_column": "sales_amount"
  }
}
```

この設定では，以下の計算を行います．

```text
sales_amount = quantity × unit_price
```

例：

| quantity | unit_price | sales_amount |
| -------: | ---------: | -----------: |
|        2 |       1800 |         3600 |
|        1 |       3200 |         3200 |
|        2 |       9600 |        19200 |

金額計算が不要な場合は，以下のように `enabled` を `false` にします．

```json
{
  "amount_calculation": {
    "enabled": false
  }
}
```

---

## 出力ファイル

## sales_summary_report.xlsx

売上明細，集計結果，エラー行をまとめて出力するExcelファイルです．

例：

```text
output/sales_summary_report.xlsx
```

このファイルには，以下のシートが含まれます．

| シート名                 | 内容                          |
| -------------------- | --------------------------- |
| `summary`            | 入力ファイル名，件数，合計金額，開始終了時刻などの概要 |
| `sales_details`      | 商品マスタと一致し，集計対象になった売上明細      |
| `monthly_sales`      | 月別の売上集計                     |
| `category_sales`     | カテゴリ別の売上集計                  |
| `product_sales`      | 商品別の売上集計                    |
| `unmatched_products` | 商品マスタに存在しなかった売上行            |
| `validation_errors`  | 売上データ自体に不備がある行              |

---

## summary シート

処理結果の概要を出力するシートです．

出力例：

| item                   | value                                                |
| ---------------------- | ---------------------------------------------------- |
| sales_files            | sample_sales_2026_06.xlsx, sample_sales_2026_07.xlsx |
| master_file            | product_master.xlsx                                  |
| report_file            | sales_summary_report.xlsx                            |
| total_files            | 2                                                    |
| total_sales_rows       | 12                                                   |
| master_rows            | 4                                                    |
| valid_sales_rows       | 8                                                    |
| sales_detail_rows      | 6                                                    |
| validation_error_rows  | 4                                                    |
| unmatched_product_rows | 2                                                    |
| total_quantity         | 13                                                   |
| total_sales            | 53550                                                |

---

## sales_details シート

商品マスタと一致した売上データを出力するシートです．

このシートに出力される行だけが，月別・カテゴリ別・商品別集計の対象になります．

出力例：

| source_file               | row_number | sales_id     | sales_date | sales_month | product_code | product_name | category | quantity | unit_price | sales_amount |
| ------------------------- | ---------: | ------------ | ---------- | ----------- | ------------ | ------------ | -------- | -------: | ---------: | -----------: |
| sample_sales_2026_06.xlsx |          2 | S-202606-001 | 2026-06-01 | 2026-06     | P-1001       | ワイヤレスマウス     | PC周辺機器   |        2 |       1800 |         3600 |
| sample_sales_2026_06.xlsx |          3 | S-202606-002 | 2026-06-02 | 2026-06     | P-1002       | USBキーボード     | PC周辺機器   |        1 |       3200 |         3200 |
| sample_sales_2026_06.xlsx |          4 | S-202606-003 | 2026-06-03 | 2026-06     | P-1004       | 外付けSSD       | ストレージ    |        1 |       9800 |         9800 |

---

## monthly_sales シート

月別の売上集計を出力するシートです．

出力例：

| sales_month | sales_count | total_quantity | total_sales |
| ----------- | ----------: | -------------: | ----------: |
| 2026-06     |           3 |              4 |       16600 |
| 2026-07     |           3 |              9 |       36950 |

---

## category_sales シート

カテゴリ別の売上集計を出力するシートです．

出力例：

| category | sales_count | total_quantity | total_sales |
| -------- | ----------: | -------------: | ----------: |
| ストレージ    |           2 |              3 |       29000 |
| PC周辺機器   |           4 |             10 |       24550 |

---

## product_sales シート

商品別の売上集計を出力するシートです．

出力例：

| product_code | product_name | category | sales_count | total_quantity | total_sales |
| ------------ | ------------ | -------- | ----------: | -------------: | ----------: |
| P-1004       | 外付けSSD       | ストレージ    |           2 |              3 |       29000 |
| P-1001       | ワイヤレスマウス     | PC周辺機器   |           2 |              7 |       12350 |
| P-1003       | Webカメラ       | PC周辺機器   |           1 |              2 |        9000 |
| P-1002       | USBキーボード     | PC周辺機器   |           1 |              1 |        3200 |

---

## unmatched_products シート

売上データとしては正常だが，商品マスタに存在しなかった行を出力するシートです．

例：

| source_file               | row_number | sales_id     | sales_date | product_code | quantity | unit_price | error_type       | error_message                    |
| ------------------------- | ---------: | ------------ | ---------- | ------------ | -------: | ---------: | ---------------- | -------------------------------- |
| sample_sales_2026_06.xlsx |          5 | S-202606-004 | 2026-06-04 | P-9999       |        3 |       1000 | master_not_found | product_code not found in master |
| sample_sales_2026_07.xlsx |          7 | S-202607-006 | 2026-07-06 | P-2000       |        1 |      12000 | master_not_found | product_code not found in master |

このシートに出力された行は，商品名やカテゴリを補完できないため，集計対象には含まれません．

---

## validation_errors シート

売上データ自体に不備がある行を出力するシートです．

例：

| source_file               | row_number | sales_id     | sales_date | product_code | quantity | unit_price | error_type       | error_message                   |
| ------------------------- | ---------: | ------------ | ---------- | ------------ | -------- | ---------- | ---------------- | ------------------------------- |
| sample_sales_2026_06.xlsx |          6 | S-202606-005 | 2026/06/05 | P-1003       | abc      | 4500       | validation_error | quantity must be numeric        |
| sample_sales_2026_06.xlsx |          7 | S-202606-006 | not-date   | P-1001       | 1        | 1800       | validation_error | sales_date must be a valid date |
| sample_sales_2026_07.xlsx |          5 | S-202607-004 | 2026/07/04 |              | 1        | 3200       | validation_error | product_code is required        |
| sample_sales_2026_07.xlsx |          6 | S-202607-005 | 2026/07/05 | P-1002       | 2        | price?     | validation_error | unit_price must be numeric      |

`source_file` と `row_number` を見れば，どのファイルの何行目を修正すればよいか確認できます．

---

## error_type の種類

| error_type         | 内容                      |
| ------------------ | ----------------------- |
| `validation_error` | 売上データ自体に不備がある           |
| `master_not_found` | 売上データの商品コードが商品マスタに存在しない |

### validation_error の例

```text
product_code is required
quantity must be numeric
unit_price must be numeric
sales_date must be a valid date
```

### master_not_found の例

```text
product_code not found in master
```

---

## 照合結果の考え方

このツールでは，売上データを以下のように分類します．

```text
売上データに問題がなく，商品マスタにも存在する行
        → sales_details シート
        → 月別・カテゴリ別・商品別集計の対象

売上データ自体に不備がある行
        → validation_errors シート
        → 集計対象外

売上データに問題はないが，商品マスタに存在しない行
        → unmatched_products シート
        → 集計対象外
```

たとえば，サンプルデータでは以下の結果になります．

```text
売上データ件数: 12
入力値検査OK: 8
集計対象: 6
入力データ不備: 4
マスタ未登録: 2
```

---

## CSV文字コードについて

CSV入力時は，以下の文字コードを順番に試します．

```text
utf-8-sig
utf-8
cp932
```

これにより，日本語を含むCSVでも読み込める可能性が高くなります．

---

## 注意事項

* 入力ファイルは `.csv` / `.xlsx` に対応しています．
* 出力ファイルは `.xlsx` 形式です．
* `.xls` 形式には対応していません．
* パスワード付きExcelや壊れたExcelファイルには対応していません．
* Excel入力ファイルは最初のシートを読み込みます．
* 入力ファイルの1行目は列名として扱います．
* 売上データは `config.json` の `sales_files` に指定した複数ファイルを処理できます．
* 売上データの照合キー列は，`sales_required_columns` に含める必要があります．
* 商品マスタの照合キー列は，`master_required_columns` に含める必要があります．
* 商品マスタの照合キーに空欄や重複がある場合，処理全体を停止します．
* 商品マスタは正しい基準データとして扱います．
* `master_value_columns` に `master_key` と同じ列を含めることはできません．
* `amount_calculation` を有効にする場合，数量列と単価列は `sales_numeric_columns` に含める必要があります．
* 商品マスタに存在しない商品コードは `unmatched_products` シートに出力され，集計対象には含まれません．
* 入力値エラー行は `validation_errors` シートに出力され，集計対象には含まれません．
* 集計結果は `sales_details` シートに出力された行をもとに作成されます．
* `output` フォルダ内の生成ファイルは実行結果なので，Git管理対象から外しても問題ありません．
* `config.json` やサンプルファイルが既に存在する場合，自動作成処理では上書きされません．
* 同名の出力ファイルが既にある場合は，実行時に上書きされます．

---

## 活用方法

* 複数売上ファイルの一括集計
* 売上CSV・Excelの自動集計
* 売上データと商品マスタの照合
* 商品コードから商品名を自動補完
* 商品カテゴリの自動付与
* 月別売上レポートの作成
* カテゴリ別売上レポートの作成
* 商品別売上レポートの作成
* 商品マスタ未登録データの抽出
* 売上データの入力ミス確認
* Excelのピボットテーブル作業の自動化
* VLOOKUP / XLOOKUP 作業の自動化
* CSV・Excel業務自動化案件のポートフォリオ

---

## Requirements

```text
pandas
openpyxl
```

---

# English

## Overview

Excel CSV Sales Summary Reporter is a Python tool that reads multiple sales CSV or Excel files，matches them with product master data，and exports a sales summary report as an Excel file.

The tool reads sales files from the `input` folder and product master data from the `master` folder.
Rows whose product codes exist in the master data are enriched with product name and category，then used for monthly，category，and product sales summaries.
Rows with validation errors or product codes not found in the master data are exported to separate sheets.

---

## Features

* Read sales CSV files
* Read sales Excel files
* Read multiple sales files
* Read product master CSV files
* Read product master Excel files
* Match sales data with product master data
* Enrich matched rows with product name and category
* Calculate sales amount from quantity and unit price
* Create monthly sales summary
* Create category sales summary
* Create product sales summary
* Extract product codes not found in master data
* Validate required fields
* Validate numeric fields
* Normalize date fields
* Check blank master keys
* Check duplicated master keys
* Export a multi-sheet Excel report
* Create sample files on first run
* Support common Japanese CSV encodings

---

## Directory Structure

```text
excel-csv-sales-summary-reporter/
├─ excel_csv_sales_summary_reporter.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  ├─ sample_sales_2026_06.xlsx
│  └─ sample_sales_2026_07.xlsx
├─ master/
│  └─ product_master.xlsx
└─ output/
   └─ sales_summary_report.xlsx
```

---

## How to Use

### 1. Install requirements

```bash
pip install -r requirements.txt
```

### 2. Place sales data

Place sales CSV or Excel files in the `input` folder.

```text
input/sample_sales_2026_06.xlsx
input/sample_sales_2026_07.xlsx
```

### 3. Place master data

Place a product master CSV or Excel file in the `master` folder.

```text
master/product_master.xlsx
```

### 4. Edit config.json

Set the sales file names，master file name，report file name，matching keys，required columns，numeric columns，date columns，and amount calculation settings.

### 5. Run the script

```bash
python excel_csv_sales_summary_reporter.py
```

---

## Output

Generated files are saved in the `output` folder.

```text
output/sales_summary_report.xlsx
```

The Excel report contains the following sheets.

| Sheet                | Description                                                  |
| -------------------- | ------------------------------------------------------------ |
| `summary`            | Summary of the processing result                             |
| `sales_details`      | Sales rows successfully matched with master data             |
| `monthly_sales`      | Monthly sales summary                                        |
| `category_sales`     | Category sales summary                                       |
| `product_sales`      | Product sales summary                                        |
| `unmatched_products` | Sales rows whose product codes were not found in master data |
| `validation_errors`  | Sales rows with validation errors                            |

---

## Result Classification

The tool classifies sales rows as follows.

```text
Rows with valid sales data and existing product codes
        -> sales_details
        -> included in summaries

Rows with validation errors
        -> validation_errors
        -> excluded from summaries

Rows with product codes not found in master data
        -> unmatched_products
        -> excluded from summaries
```

---

## Limitations

* Input files must be `.csv` or `.xlsx`.
* Output files are `.xlsx`.
* `.xls` files are not supported.
* Password-protected or broken Excel files are not supported.
* For Excel input，only the first sheet is read.
* The master key must not be blank.
* The master key must not be duplicated.
* Rows in `validation_errors` are excluded from summaries.
* Rows in `unmatched_products` are excluded from summaries.

---

## Use Cases

* Summarize multiple sales files
* Match sales data with product master data
* Enrich product names and categories
* Extract unknown product codes
* Validate sales data
* Create monthly sales reports
* Create category sales reports
* Create product sales reports
* Automate VLOOKUP-like Excel tasks
* Automate CSV / Excel sales reporting workflows
