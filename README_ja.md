
# ログ比較ツール 使用方法

このリポジトリは、STG 環境と PRD 環境で取得したログを比較し、差分を Excel 形式で出力するためのスクリプト群を提供します。

## 前提条件

- Python 3.10 以上
- `ref/` ディレクトリ直下に STG/PRD ログファイルが配置されていること  
  - ファイル名フォーマット例：`20251028_1035_GetInfo_OS_diff_event-lb01s.log`
  - 対応する PRD ファイル（末尾 `p`）が存在する必要があります

## 主要スクリプト

- `log_to_excel.py`  
  ログから差分表を生成し Excel ファイルに出力します。
- `verify_workbook.py`  
  生成された Excel にすべてのログ行が含まれているか検証します。

## 使い方

### 1. ログ比較結果の生成

```bash
python log_to_excel.py --input-dir ref --output comparison.xlsx
```

- `--input-dir`: ログが置かれているディレクトリ（省略時はカレントディレクトリ）
- `--output`: 生成する Excel ファイルのパス（既定値は `comparison.xlsx`）
- 実行すると、対象ディレクトリ内の STG/PRD ログをペアリングし、1 ペアごとにシートを作成して差分を出力します。
- 既存の Excel に同名シートがある場合は `.v2`, `.v3` のようにサフィックスを付けて追記します。

### 2. ログの検証

```bash
python verify_workbook.py
```

- `ref/` 内のログと `comparison.xlsx` を照合し、欠落行がないか確認します。
- 別ディレクトリや別 Excel を検証したい場合は、`verify_logs_against_workbook(Path(...), Path(...))` を直接呼び出します。

## 出力される Excel の形式

各シートの先頭には、リファレンス Excel (`ref/` 以下の期待値ファイル) に倣ったヘッダーが生成されます。

- A1: `InfoOne延命プロジェクト`
- A2: `オンプレ単体テスト（NewSTG基盤）`
- B4: `対象サーバ`
- C4: `備考`
- B5: `=RIGHT(CELL("filename",A1),LEN(CELL("filename",A1))-FIND("]",CELL("filename",A1)))`  
  （シート名からサーバ名を取得する式）
- 7 行目: `現行サーバ（xxx） / 差分有無 / 新基盤（xxx） / 備考`
- 8 行目以降: STG/PRD ログの行を並べ、C 列には `B列 = D列` を評価する式を設定

## 典型的なワークフロー

1. `ref/` に STG/PRD ログを配置する。
2. `python log_to_excel.py --input-dir ref --output comparison.xlsx` を実行。
3. `python verify_workbook.py` で取り込み漏れがないかチェック。
4. 必要に応じて `comparison.xlsx` をレビュー・共有。

## 補足

- `single_ref/` や `test/` は動作確認用途のディレクトリです。`.gitignore` で除外しています。
- 生成物は自動的に上書きされます。履歴が必要な場合は適宜バックアップを取ってください。
