---
name: excel-markdown-workflow
description: >-
  markitdownで生成されたExcelマークダウンファイルの整形ワークフローを案内するスキル。
  ユーザーがExcelをマークダウンに変換したい、markitdownの出力を整形したい、
  Excel方眼紙を処理したいと言及した際に自動トリガーされる。
---

# Excel Markdown Reformatter ワークフロー

markitdownで生成されたExcelマークダウンファイルを、読みやすい形式に再フォーマットするパイプラインです。

## 前提条件

uv（Pythonパッケージマネージャ）が必要です。markitdownはuvx経由で一時実行されるため、事前インストール不要です。

```bash
# uvのインストール（未導入の場合）
curl -LsSf https://astral.sh/uv/install.sh | sh
```

## 推奨ワークフロー（ワンショット実行）

.xlsxファイルを渡すだけで全工程を自動実行します。

```bash
# 単一ファイル
/excel-md:convert your_excel_file.xlsx

# 複数ファイル
/excel-md:convert file1.xlsx file2.xlsx

# ディレクトリ内の全xlsx
/excel-md:convert ./data/
```

markitdown変換 → prepare → transform（並列） → merge の全工程が自動実行され、`{basename}_reformed.md` が生成されます。transform処理ではsheet-transformerサブエージェントにより、最大5シート同時の並列変換が行われます。

## カスタム/デバッグ用ワークフロー（個別ステップ実行）

特定ステップだけ再実行したい場合や、変換パラメータを調整したい場合に使用します。

### Step 1: Excel → Markdown変換（外部ツール）

```bash
uvx markitdown your_excel_file.xlsx > your_excel_file.md
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

sheet-transformerサブエージェント並列実行で各シートを個別のマークダウンファイルに変換します（5シートごとのバッチ処理）。

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

- 並列処理: sheet-transformerサブエージェントによる5シートごとの分割バッチ処理
- 1シート=1サブエージェント方式で各シートを独立処理
- 見出しレベル制御: シート名は ## レベル（h2）、シート内見出しは ### レベル（h3）以下
