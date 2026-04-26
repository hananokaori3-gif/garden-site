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
import io
from pathlib import Path

import anthropic
import pandas as pd
from pdf2image import convert_from_path


# ──────────────────────────────────────────
# 1. PDF → Claude API でデータ抽出
# ──────────────────────────────────────────

def image_to_base64(image):
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode()


def extract_from_pdf(client: anthropic.Anthropic, pdf_path: str) -> list[dict]:
    """画像PDFの各ページをClaude APIで読み取り、品名・金額を抽出"""
    print(f"PDFを画像に変換中: {pdf_path}")
    pages = convert_from_path(pdf_path, dpi=200)
    print(f"  {len(pages)} ページ検出")

    all_items = []
    for i, page in enumerate(pages, 1):
        print(f"  ページ {i}/{len(pages)} を読み取り中...")
        img_b64 = image_to_base64(page)

        prompt = """この仕入れ伝票の明細行をすべて読み取ってください。

以下のJSON形式のみで返してください（説明不要）:
[
  {"品名": "商品名", "金額": 数値},
  ...
]

注意:
- 金額は税抜き金額の数値のみ（カンマ・円記号は除く）
- 品名が空白・合計行・小計行・消費税行はスキップ
- 読み取れない場合は [] を返す"""

        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {
                        "type": "base64", "media_type": "image/png", "data": img_b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )

        text = message.content[0].text.strip()
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
    """奉行クラウド出力Excelを読み込む（商品名・税抜純仕入高）"""
    for skip in range(0, 10):
        df = pd.read_excel(excel_path, skiprows=skip)
        df.columns = df.columns.astype(str).str.strip()
        name_col = find_col(df, ["商品名", "品名", "品目", "摘要", "商品コード名"])
        amt_col  = find_col(df, ["税抜純仕入高", "税抜き純仕入れ高", "税抜仕入高",
                                  "税抜金額", "仕入金額", "仕入額", "金額", "金額(円)"])
        if name_col and amt_col:
            df = df[[name_col, amt_col]].copy()
            df.columns = ["商品名", "税抜純仕入高"]
            df["商品名"] = df["商品名"].astype(str).str.strip()
            df = df[df["商品名"].notna() & (df["商品名"] != "") & (df["商品名"] != "nan")]
            df["税抜純仕入高"] = pd.to_numeric(df["税抜純仕入高"], errors="coerce")
            return df.dropna(subset=["税抜純仕入高"])

    print("エラー: Excelに「商品名」「税抜純仕入高」の列が見つかりません。")
    print(f"  列名一覧: {list(df.columns)}")
    sys.exit(1)


# ──────────────────────────────────────────
# 3. 品名の名寄せ（Claude APIによる意味的マッチング）
# ──────────────────────────────────────────

def match_names(client: anthropic.Anthropic,
                pdf_names: list[str], excel_names: list[str]) -> list[dict]:
    """
    PDF側の品名リストと奉行側の商品名リストを渡し、
    同一商品と思われるペアをClaude APIに対応付けさせる。
    """
    print("\n品名の名寄せ中（表記ゆれを自動対応）...")

    prompt = f"""以下は仕入れ伝票（PDF）の品名リストと、奉行クラウド（Excel）の商品名リストです。
表記が違っていても同じ商品と判断できるものを対応付けてください。

【PDF品名リスト】
{json.dumps(pdf_names, ensure_ascii=False, indent=2)}

【奉行商品名リスト】
{json.dumps(excel_names, ensure_ascii=False, indent=2)}

以下のJSON形式のみで返してください（説明不要）:
[
  {{"pdf品名": "PDF側の品名", "奉行商品名": "奉行側の商品名", "信頼度": "高/中/低"}},
  ...
]

ルール:
- 対応できるものだけ列挙する（無理に対応付けない）
- 片方にしか存在しない商品は含めない
- 表記ゆれの例: ひらがな↔カタカナ、漢字↔かな、略称、サイズ表記の違い など
- 信頼度: 確実に同じなら「高」、おそらく同じなら「中」、念のため確認が必要なら「低」"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    text = message.content[0].text.strip()
    if "```" in text:
        text = text.split("```")[1].replace("json", "").strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        print(f"  ⚠ 名寄せ結果の解析に失敗しました。出力: {text[:200]}")
        return []


# ──────────────────────────────────────────
# 4. 突き合わせ
# ──────────────────────────────────────────

def compare(client: anthropic.Anthropic,
            pdf_items: list[dict], excel_df: pd.DataFrame):

    pdf_df = pd.DataFrame(pdf_items)[["品名", "金額"]].copy()
    pdf_df["品名"] = pdf_df["品名"].astype(str).str.strip()
    pdf_df["金額"] = pd.to_numeric(pdf_df["金額"], errors="coerce")

    # 名寄せ
    pdf_names   = pdf_df["品名"].unique().tolist()
    excel_names = excel_df["商品名"].unique().tolist()
    matches     = match_names(client, pdf_names, excel_names)

    # マッチング結果を表示
    print(f"\n  対応付け結果: {len(matches)} 件")
    for m in matches:
        mark = {"高": "✓", "中": "△", "低": "?"}.get(m.get("信頼度", ""), "")
        print(f"  {mark} [{m.get('信頼度','')}] 「{m['pdf品名']}」→「{m['奉行商品名']}」")

    # マッピング辞書（PDF品名 → 奉行商品名）
    name_map = {m["pdf品名"]: m["奉行商品名"] for m in matches}

    # PDF側に「対応する奉行商品名」列を付加
    pdf_df["奉行商品名"] = pdf_df["品名"].map(name_map)

    issues = []

    # マッチした行を金額比較
    matched_excel_names = set()
    for _, row in pdf_df.iterrows():
        pdf_name   = row["品名"]
        excel_name = row["奉行商品名"]
        amt_p      = row["金額"]

        if pd.isna(excel_name):
            # 奉行に対応商品なし
            issues.append({
                "PDF品名": pdf_name, "奉行商品名": "-",
                "種別": "伝票のみ（奉行に未入力 or 名寄せ不可）",
                "金額_伝票": amt_p, "金額_奉行": "-", "差額": "", "信頼度": ""
            })
            continue

        matched_excel_names.add(excel_name)
        excel_row = excel_df[excel_df["商品名"] == excel_name]
        if excel_row.empty:
            continue
        amt_e      = excel_row["税抜純仕入高"].values[0]
        confidence = next((m["信頼度"] for m in matches if m["奉行商品名"] == excel_name), "")

        if abs(amt_p - amt_e) > 1:
            diff = amt_p - amt_e
            issues.append({
                "PDF品名": pdf_name, "奉行商品名": excel_name,
                "種別": "金額が不一致",
                "金額_伝票": amt_p, "金額_奉行": amt_e,
                "差額": f"{diff:+,.0f}", "信頼度": confidence
            })

    # 奉行にあってPDFで対応付けされなかった行
    for _, row in excel_df.iterrows():
        if row["商品名"] not in matched_excel_names:
            issues.append({
                "PDF品名": "-", "奉行商品名": row["商品名"],
                "種別": "奉行のみ（伝票に未記載 or 名寄せ不可）",
                "金額_伝票": "-", "金額_奉行": row["税抜純仕入高"],
                "差額": "", "信頼度": ""
            })

    total_p = pdf_df["金額"].sum()
    total_e = excel_df["税抜純仕入高"].sum()

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
            conf = f"  ［信頼度:{iss['信頼度']}］" if iss["信頼度"] else ""
            print(f"  [{i}] {iss['種別']}{conf}")
            print(f"       伝票品名:   {iss['PDF品名']}")
            print(f"       奉行商品名: {iss['奉行商品名']}")
            if iss["差額"]:
                print(f"       金額:  伝票={iss['金額_伝票']:,.0f}  /  奉行={iss['金額_奉行']:,.0f}"
                      f"  （差額 {iss['差額']} 円）")
            print()

    # Excel出力
    out = "突き合わせ結果.xlsx"
    issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(
        columns=["PDF品名", "奉行商品名", "種別", "金額_伝票", "金額_奉行", "差額", "信頼度"])
    summary_df = pd.DataFrame([
        {"項目": "仕入伝票（PDF）合計", "金額": total_p, "行数": len(pdf_df)},
        {"項目": "奉行クラウド合計",    "金額": total_e, "行数": len(excel_df)},
        {"項目": "差額",               "金額": total_p - total_e, "行数": ""},
    ])
    matches_df  = pd.DataFrame(matches) if matches else pd.DataFrame(
        columns=["pdf品名", "奉行商品名", "信頼度"])
    pdf_raw_df  = pd.DataFrame(pdf_items)

    with pd.ExcelWriter(out) as writer:
        issues_df.to_excel(writer, sheet_name="差異一覧", index=False)
        summary_df.to_excel(writer, sheet_name="合計サマリ", index=False)
        matches_df.to_excel(writer, sheet_name="品名対応表", index=False)
        pdf_raw_df.to_excel(writer, sheet_name="PDF読取データ", index=False)
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

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("エラー: ANTHROPIC_API_KEY が設定されていません。")
        print("  export ANTHROPIC_API_KEY=sk-ant-xxxx を実行してから再試行してください。")
        sys.exit(1)

    client = anthropic.Anthropic()

    print("=== 仕入れ突き合わせツール ===")
    print(f"PDF:   {pdf_path}")
    print(f"Excel: {excel_path}\n")

    pdf_items = extract_from_pdf(client, pdf_path)
    print(f"\nPDF から {len(pdf_items)} 件の明細を読み取りました。")

    excel_df = load_bugyo_excel(excel_path)
    print(f"奉行Excel から {len(excel_df)} 件の明細を読み込みました。")

    compare(client, pdf_items, excel_df)


if __name__ == "__main__":
    main()
