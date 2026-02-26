---
description: Excel ファイルをワンショットで Markdown に変換
argument-hint: "<file.xlsx|dir/|*.xlsx> [file2.xlsx ...]"
---

# Excel Markdown ワンショット変換

.xlsxファイルを受け取り、markitdown変換 → prepare → transform → merge の全工程を自動実行します。

## Phase 0: 前提条件チェック

### uv存在確認
```bash
command -v uv
```
- 成功時: 続行
- 失敗時: 以下のエラーメッセージを表示して処理を中止
  ```
  エラー: uvが見つかりません。以下のコマンドでインストールしてください:
  curl -LsSf https://astral.sh/uv/install.sh | sh
  ```

### python3存在確認
```bash
command -v python3
```
- 失敗時: エラーメッセージを表示して処理を中止

## Phase 1: 入力解析

$ARGUMENTSを解析し、処理対象のxlsxファイルリストを構築する。

### 解析ルール
- 個別ファイル指定: `file.xlsx file2.xlsx` → そのまま使用
- ディレクトリ指定: `dir/` → `dir/*.xlsx` をGlobで検索
- ワイルドカード: `*.xlsx` → Globで展開
- 各ファイルの存在確認を実行し、不在ファイルは警告を表示してスキップ

### ファイルリスト表示
対象ファイル一覧を表示し、処理を開始する。

## Phase 2: markitdown変換（順次実行）

各xlsxファイルについてBashで順次実行:

```bash
uvx --with 'markitdown[xlsx]' markitdown "{filepath}" > "{basename}.md"
```

- `{basename}` はファイル名から `.xlsx` 拡張子を除いたもの
- 失敗時はエラーを表示し、そのファイルをスキップして次のファイルへ続行
- 成功時は生成された `.md` ファイルのパスを記録

## Phase 3: ファイルごとの変換パイプライン

Phase 2で生成された各 `.md` ファイルについて、以下の3ステップを順次実行する。
複数ファイルがある場合も、ファイルごとに順次処理する。

### 3-1: Prepare（シート構造解析・変換戦略決定）

以下の処理をインラインで実行する（prepare.md相当）:

1. Readツールで `{basename}.md` を読み込む
2. `##` レベルヘッダーでシート境界を識別し、シート一覧を抽出
3. 各シートについて分析:
   - シート名とデータ範囲の特定
   - NaN値とUnnamed列のカウント
   - テーブル構造の複雑度判定
4. シート数による並列処理方式の判定（≤5: 単一バッチ、>5: 5シートごと分割バッチ）
5. `{basename}_transform_prompt.md` として変換プロンプトを保存
   - 各シートの個別処理指示（1シート=1Task方式）
   - NaN/Unnamed除去パターン
   - ファイル命名規則: `{basename}_sheet_{2桁番号}_{シート名}.md`
   - 厳格禁止事項と必須実行事項
   - 見出しレベル制御: シート名は ## レベル、シート内見出しは ### 以下

### 3-2: Transform（sheet-transformerサブエージェント並列起動）

#### シートデータの抽出
`{basename}.md` を `##` レベルヘッダーで分割し、各シートのraw markdownを抽出する。

#### sheet-transformerサブエージェント並列起動

以下の必須ルールに従って並列実行する:

##### 並列起動の必須ルール
- 各シートに対して1つのTask Tool呼び出しを作成する
- `subagent_type: "sheet-transformer"` を指定する
- 依存関係のない複数のTask Tool呼び出しは、必ず同一ターン内で並列発行すること
- シート数≤5: 全シートを1回の並列発行で実行
- シート数>5: 5シートずつバッチ化し、バッチ完了後に次バッチを発行

##### 各サブエージェントへのプロンプト構成
promptには以下を全て含む自己完結したテキストを渡す:
1. 「以下のシートを変換してください」という明確な指示
2. 出力ファイルパス: `{basename}_sheet_{2桁番号}_{シート名}.md`
3. シート名
4. 対象シートのraw markdownテキスト全文（`##` ヘッダーからの抽出済み）

##### プロンプトテンプレート
各サブエージェントに渡すプロンプトは以下の形式:
```
以下のシートを変換してください。

出力ファイルパス: {output_path}
シート名: {sheet_name}

変換対象のraw markdownテキスト:
---
{raw_markdown_content}
---
```

##### バッチ実行パターン
```
シート数≤5の場合:
  → 全シートのTask Tool呼び出しを同一メッセージで発行

シート数>5の場合:
  バッチ1: シート1-5のTask Tool呼び出しを同一メッセージで発行
  → バッチ1の全Task完了を待機
  バッチ2: シート6-10のTask Tool呼び出しを同一メッセージで発行
  → バッチ2の全Task完了を待機
  ...全シート完了まで繰り返し
```

#### 完了確認
全サブエージェント完了後:
- 生成されたファイル一覧をGlobで確認: `{basename}_sheet_*.md`
- 欠落シートがあれば警告を表示

### 3-3: Merge（統合処理）

Pythonスクリプトで統合を実行する。

#### merge_sheets.py パス解決
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/merge_sheets.py" "{basename}"
```

- 成功時: `{basename}_reformed.md` が生成される
- 失敗時: エラーを表示し、手動mergeの指示を出す

## Phase 4: 最終サマリー

全ファイルの処理完了後、以下の情報を表示:

### 処理結果
- 成功ファイル一覧と出力パス（`{basename}_reformed.md`）
- 失敗ファイル一覧と失敗理由
- 各ファイルのシート数と処理統計

### 生成ファイル
- 最終出力: `{basename}_reformed.md`
- 中間ファイル: `{basename}_transform_prompt.md`, `{basename}_sheet_*.md`

## エラー処理

- Phase 2でmarkitdown失敗: そのファイルをスキップし、次のファイルへ
- Phase 3-2でサブエージェント失敗: 失敗シートを報告し、手動リトライの指示を出す
- Phase 3-3でmerge失敗: `/excel-md:merge` の手動実行を案内
- 全Phase共通: 致命的エラー以外は処理を継続し、最終サマリーで全エラーを集約表示

## 実行例

```bash
# 単一ファイル
/excel-md:convert report.xlsx

# 複数ファイル
/excel-md:convert file1.xlsx file2.xlsx

# ディレクトリ内の全xlsx
/excel-md:convert ./data/

# ワイルドカード
/excel-md:convert *.xlsx
```
