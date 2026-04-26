"""
仕入れ伝票（画像PDF） × 奉行クラウド（Excel）突き合わせツール

使い方:
  python3 仕入突き合わせ.py 仕入伝票.pdf 奉行出力.xlsx

必要な環境変数:
  ANTHROPIC_API_KEY=sk-ant-...  （Claude APIキー）

インストール:
  pip install anthropic pdf2image pillow pandas openpyxl
  apt install poppler-utils  （Mac: brew install poppler）
"""

import os
import sys
import json
import base64
import tempfile
from pathlib import Path

import anthropic
import pandas as pd
from pdf2image import convert_from_path


# ──────────────────────────────────────────
# 1. PDF → Claude API でデータ抽出
# ──────────────────────────────────────────

def image_to_base64(image):
    import io
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode()


def extract_from_pdf(pdf_path: str) -> list[dict]:
    """画像PDFの各ページをClaude APIで読み取り、品名・数量・金額を抽出"""
    client = anthropic.Anthropic()  # ANTHROPIC_API_KEY を自動読み込み

    print(f"PDFを画像に変換中: {pdf_path}")
    pages = convert_from_path(pdf_path, dpi=200)
    print(f"  {len(pages)} ページ検出")

    all_items = []

    for i, page in enumerate(pages, 1):
        print(f"  ページ {i}/{len(pages)} を読み取り中...")
        img_b64 = image_to_base64(page)

        prompt = """この仕入れ伝票から、すべての明細行を読み取ってください。

以下のJSON形式で返してください（余計な説明は不要、JSONのみ）:
[
  {"品名": "商品名", "金額": 数値},
  ...
]

注意:
- 金額は税抜き金額の数値のみ（カンマや円記号は除く）
- 品名が空白・合計行・小計行はスキップ
- 読み取れない場合は空リスト [] を返す"""

        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": img_b64,
                        },
                    },
                    {"type": "text", "text": prompt},
                ],
            }],
        )

        text = message.content[0].text.strip()
        # JSON部分だけ取り出す
        if "```" in text:
            text = text.split("```")[1].replace("json", "").strip()

        try:
            items = json.loads(text)
            for item in items:
                item["ページ"] = i
            all_items.extend(items)
        except json.JSONDecodeError:
            print(f"  ⚠ ページ {i} の解析に失敗しました。出力: {text[:100]}")

    return all_items


# ──────────────────────────────────────────
# 2. 奉行クラウド Excel 読み込み
# ──────────────────────────────────────────

def find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def load_bugyo_excel(excel_path: str) -> pd.DataFrame:
    """奉行クラウド出力Excelを読み込む"""
    # ヘッダー行を自動検出（最初の数行をスキャン）
    for skip in range(0, 10):
        df = pd.read_excel(excel_path, skiprows=skip)
        df.columns = df.columns.astype(str).str.strip()
        name_col = find_col(df, ["商品名", "品名", "品目", "摘要", "商品コード名"])
        amt_col  = find_col(df, ["税抜純仕入高", "税抜き順仕入れ高", "税抜順仕入高",
                                  "税抜仕入高", "税抜金額", "仕入金額", "仕入額", "金額", "金額(円)"])
        if name_col and amt_col:
            df = df[[name_col, amt_col]].copy()
            df.columns = ["品名", "金額"]
            df["品名"] = df["品名"].astype(str).str.strip()
            df = df[df["品名"].notna() & (df["品名"] != "") & (df["品名"] != "nan")]
            df["金額"] = pd.to_numeric(df["金額"], errors="coerce")
            return df.dropna(subset=["金額"])

    print("エラー: Excelに「商品名」「税抜き順仕入れ高」の列が見つかりません。")
    print(f"  列名一覧: {list(df.columns)}")
    print("  スクリプト内の find_col() の候補リストに列名を追加してください。")
    sys.exit(1)


# ──────────────────────────────────────────
# 3. 突き合わせ
# ──────────────────────────────────────────

def compare(pdf_items: list[dict], excel_df: pd.DataFrame):
    pdf_df = pd.DataFrame(pdf_items)[["品名", "金額"]].copy()
    pdf_df["品名"] = pdf_df["品名"].astype(str).str.strip()
    pdf_df["金額"] = pd.to_numeric(pdf_df["金額"], errors="coerce")

    merged = pd.merge(pdf_df, excel_df, on="品名", how="outer",
                      suffixes=("_伝票", "_奉行"))

    issues = []
    for _, row in merged.iterrows():
        name  = row["品名"]
        amt_p = row.get("金額_伝票")
        amt_e = row.get("金額_奉行")

        if pd.isna(amt_p):
            issues.append({"品名": name, "種別": "奉行のみ（伝票に未記載？）",
                            "金額_伝票": "-", "金額_奉行": amt_e, "差額": ""})
        elif pd.isna(amt_e):
            issues.append({"品名": name, "種別": "伝票のみ（奉行に未入力）",
                            "金額_伝票": amt_p, "金額_奉行": "-", "差額": ""})
        elif abs(amt_p - amt_e) > 1:
            diff = amt_p - amt_e
            issues.append({"品名": name, "種別": "金額が不一致",
                            "金額_伝票": amt_p, "金額_奉行": amt_e,
                            "差額": f"{diff:+,.0f}"})

    total_p = pdf_df["金額"].sum()
    total_e = excel_df["金額"].sum()

    # 画面表示
    print("\n" + "=" * 65)
    print("  突き合わせ結果")
    print("=" * 65)
    print(f"  仕入伝票（PDF）合計:  {total_p:>12,.0f} 円  ({len(pdf_df)} 行)")
    print(f"  奉行クラウド合計:     {total_e:>12,.0f} 円  ({len(excel_df)} 行)")
    print(f"  差額:                 {total_p - total_e:>+12,.0f} 円")
    print("=" * 65)

    if not issues:
        print("  差異なし：すべて一致しています。")
    else:
        print(f"  差異 {len(issues)} 件\n")
        for i, iss in enumerate(issues, 1):
            print(f"  [{i}] {iss['種別']}")
            print(f"       品名:  {iss['品名']}")
            print(f"       金額:  伝票={iss['金額_伝票']}  /  奉行={iss['金額_奉行']}"
                  + (f"  （差額 {iss['差額']} 円）" if iss['差額'] else ""))
            print()

    # Excel出力
    out = "突き合わせ結果.xlsx"
    issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(
        columns=["品名", "種別", "金額_伝票", "金額_奉行", "差額"])
    summary_df = pd.DataFrame([
        {"項目": "仕入伝票（PDF）合計", "金額": total_p, "行数": len(pdf_df)},
        {"項目": "奉行クラウド合計",    "金額": total_e, "行数": len(excel_df)},
        {"項目": "差額",               "金額": total_p - total_e, "行数": ""},
    ])
    pdf_detail_df = pd.DataFrame(pdf_items)
    with pd.ExcelWriter(out) as writer:
        issues_df.to_excel(writer, sheet_name="差異一覧", index=False)
        summary_df.to_excel(writer, sheet_name="合計サマリ", index=False)
        pdf_detail_df.to_excel(writer, sheet_name="PDF読取データ", index=False)
        excel_df.to_excel(writer, sheet_name="奉行データ", index=False)

    print(f"  結果を保存しました: {out}")
    return issues


# ──────────────────────────────────────────
# メイン
# ──────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    pdf_path   = sys.argv[1]
    excel_path = sys.argv[2]

    if not Path(pdf_path).exists():
        print(f"エラー: PDFファイルが見つかりません: {pdf_path}")
        sys.exit(1)
    if not Path(excel_path).exists():
        print(f"エラー: Excelファイルが見つかりません: {excel_path}")
        sys.exit(1)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("エラー: ANTHROPIC_API_KEY が設定されていません。")
        print("  export ANTHROPIC_API_KEY=sk-ant-xxxx を実行してから再試行してください。")
        sys.exit(1)

    print("=== 仕入れ突き合わせツール ===")
    print(f"PDF:   {pdf_path}")
    print(f"Excel: {excel_path}\n")

    pdf_items = extract_from_pdf(pdf_path)
    print(f"\nPDF から {len(pdf_items)} 件の明細を読み取りました。\n")

    excel_df = load_bugyo_excel(excel_path)
    print(f"奉行Excel から {len(excel_df)} 件の明細を読み込みました。\n")

    compare(pdf_items, excel_df)


if __name__ == "__main__":
    main()
