
# 内部実装メモ

## 全体構成

- `log_to_excel.py`
  - STG/PRD ログのペア探索 (`_discover_pairs`)
  - 行単位の差分整列 (`difflib.SequenceMatcher`)
  - 既存 Excel の読み込み・追記 (`_load_existing_sheets`)
- `xlsx_writer.py`
  - 依存ライブラリ無しで XLSX を生成
  - `CellValue` で値・式・型を保持し `<c>`, `<f>`, `<v>` を出力
- `verify_workbook.py`
  - 生成済みシートのテキストを抽出し、ログ行がすべて存在するか検証

## ログ整列

- `SequenceMatcher` の `get_opcodes()` を利用し、`equal` / `replace` / `delete` / `insert` を処理。
- 行番号は `S:`（STG）と `P:`（PRD）で保持。
- C列には `B列 = D列` を評価する式を付与 (`CellValue(formula=..., data_type="b")`)。

## Excel 追記処理

- 既存ファイルがある場合は `zipfile` + `ElementTree` で `xl/workbook.xml` と `xl/_rels/workbook.xml.rels` を解析。
- シート XML を `CellValue` 配列に変換し、元の式／値を保持する。
- 同名シートがある場合は `.v2`, `.v3` … を付与して衝突を回避。

## スタイル

- `styles.xml` は最小構成（Calibri 11pt）を固定で出力。
- 先頭7行のヘッダーは期待値ファイルを模倣。B5 のサーバ名は `RIGHT(CELL("filename"...))` で動的取得。

## 留意事項

- セルの `data_type` は `None`（inlineStr）、`"b"`（論理値）、`"str"`（式結果）を使用。
- 既存ブックの sharedStrings を読み込み、セル参照 `s` の値を文字列に復元。
- 参照行が多いため生成結果は数 MB 程度になる。必要に応じて gzip 等で圧縮可能。
