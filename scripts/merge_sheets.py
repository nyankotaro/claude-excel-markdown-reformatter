#!/usr/bin/env python3
"""
Excel Markdown Reformatter - Sheet Merger
シートファイルを統合して reformed.md ファイルを生成
"""
import sys
import glob
from datetime import datetime

def merge_sheets(basename):
    """シートファイルを統合して reformed.md を生成"""
    # シートファイルを番号順で取得
    sheet_files = sorted(glob.glob(f"{basename}_sheet_*.md"))

    if not sheet_files:
        print(f"エラー: {basename}_sheet_*.md ファイルが見つかりません")
        sys.exit(1)

    output_file = f"{basename}_reformed.md"

    # 出力ファイルに書き込み
    with open(output_file, 'w', encoding='utf-8') as outfile:
        # メタデータヘッダー
        outfile.write(f"<!-- Excel Markdown Reformatter で生成 -->\n")
        outfile.write(f"<!-- 元ファイル: {basename}.md -->\n")
        outfile.write(f"<!-- 処理日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -->\n")
        outfile.write(f"<!-- シート数: {len(sheet_files)} -->\n\n")

        # 各シートファイルを順次結合
        for sheet_file in sheet_files:
            with open(sheet_file, 'r', encoding='utf-8') as infile:
                content = infile.read().strip()
                if content:
                    outfile.write(content)
                    outfile.write("\n\n")

    print(f"✓ {len(sheet_files)}個のシートを統合 → {output_file}")

def main():
    if len(sys.argv) != 2:
        print("使用方法: python merge_sheets.py <basename>")
        print("例: python merge_sheets.py \"your_excel_file\"")
        sys.exit(1)

    basename = sys.argv[1]
    merge_sheets(basename)

if __name__ == "__main__":
    main()
