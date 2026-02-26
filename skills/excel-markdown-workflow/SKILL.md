---
name: excel-markdown-workflow
description: >-
  markitdownで生成されたExcelマークダウンファイルの整形ワークフローを案内するスキル。
  ユーザーがExcelをマークダウンに変換したい、markitdownの出力を整形したい、
  Excel方眼紙を処理したいと言及した際に自動トリガーされる。
---

# Excel Markdown Reformatter ワークフロー

markitdownで生成されたExcelマークダウンファイルを、読みやすい形式に再フォーマットする3ステップのパイプラインです。

## 前提条件

markitdownライブラリが必要です。

```bash
pip install markitdown
```

## ワークフロー

### Step 1: Excel → Markdown変換（外部ツール）

```bash
markitdown your_excel_file.xlsx > your_excel_file.md
```

### Step 2: 準備と分析

```bash
/excel-md:prepare your_excel_file.md
```

ファイルを解析し、シート構造を識別して変換プロンプトを生成します。

### Step 3: シート変換

```bash
/excel-md:transform your_excel_file_transform_prompt.md
```

Task Tool並列実行で各シートを個別のマークダウンファイルに変換します（5シートごとのバッチ処理）。

### Step 4: ファイル統合

```bash
# Pythonスクリプト版（推奨、大きなファイルに対応）
python "${CLAUDE_PLUGIN_ROOT}/scripts/merge_sheets.py" "your_excel_file"
# リポジトリを直接使用している場合: python scripts/merge_sheets.py "your_excel_file"

# スラッシュコマンド版
/excel-md:merge your_excel_file
```

変換されたシートファイルを統合し、最終ドキュメント `your_excel_file_reformed.md` を生成します。

## ファイル命名規則

- 入力: `{basename}.md`（markitdown生成）
- 変換プロンプト: `{basename}_transform_prompt.md`
- 個別シート: `{basename}_sheet_{2桁番号}_{シート名}.md`
- 統合ファイル: `{basename}_reformed.md`

## データ保持原則

このツールは厳格なデータ保持原則に基づいて動作します。

- 元ファイルの各セル値を一文字も変更しない
- 推測による情報補完を完全禁止
- 構造整形とマークダウンテーブル生成のみ許可
- NaN値とUnnamed列の除去のみが許可される変更

詳細は [references/data-constraints.md](references/data-constraints.md) を参照してください。

## 処理方式

- 並列処理: 5シートごとの分割バッチ処理（Task Tool使用）
- 1シート=1Task方式で各シートを独立処理
- 見出しレベル制御: シート名は ## レベル（h2）、シート内見出しは ### レベル（h3）以下
