# Excel Markdown Reformatter

Excel方眼紙ファイルをMarkdownに変換し、読みやすい形式に再フォーマットするClaude Codeスラッシュコマンドセット

## Requirements

- markitdownライブラリ（Excelファイルの事前変換用）
- Claude Code環境

## 処理フロー

```mermaid
graph TD
    A[Excel File] -->|markitdown| B[Raw Markdown]
    B -->|/excel-md:prepare| C[Analysis & Strategy]
    C --> D[Transform Prompt]
    D -->|/excel-md:transform| E[Individual Sheet Files]
    E -->|/excel-md:merge| F[Reformed Markdown]

    style A fill:#e1f5fe
    style B fill:#fff3e0
    style F fill:#e8f5e8
```

## Commands

### `/excel-md:prepare`
Excelファイルの準備と分析を行います。

### `/excel-md:transform`
個別シートを読みやすいMarkdownに変換します。

### `/excel-md:merge`
変換されたMarkdownファイルを統合し、最終的なドキュメントを生成します。

## Workflow

### Step 1: Excel → Markdown変換
```bash
# markitdownでExcelファイルをMarkdownに変換
markitdown your_excel_file.xlsx > your_excel_file.md
```

### Step 2: ファイル準備と分析
```bash
/excel-md:prepare your_excel_file.md
```
- ファイルの検証とシート構造解析
- 変換戦略の決定と並列処理方式の判定
- 変換用プロンプトファイル生成

### Step 3: シート変換
```bash
/excel-md:transform your_excel_file_transform_prompt.md
```
- 各シートを個別のMarkdownファイルに変換
- NaN値とUnnamed列の除去
- テーブル構造の整形と品質向上

### Step 4: ファイル統合

#### Pythonスクリプト版（推奨）
```bash
python merge_sheets.py "your_excel_file"
```

**⚠️ 大きなファイルの統合について**

統合対象のシートファイルが大きい場合、スラッシュコマンド版では「Claude's response exceeded the 32000 output token maximum.」エラーが発生する可能性があります。大量のシートや大きなデータを統合する場合は、**Pythonスクリプト版の使用を推奨**します。

#### スラッシュコマンド版
```bash
/excel-md:merge your_excel_file
```

共通機能:
- 変換されたシートファイル（`*_sheet_*.md`）を統合
- 最終ドキュメント（`*_reformed.md`）生成

## Output

- 変換用プロンプトファイル
- 各シート用の個別Markdownファイル
- 統合された最終ドキュメント
