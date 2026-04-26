"""
Excel データ突き合わせツール
品名・数量・金額を比較して差異を出力します
"""

import pandas as pd
import sys
from pathlib import Path


def load_excel(path, sheet=0):
    """Excelファイルを読み込み、列名を正規化する"""
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = df.columns.str.strip()
    return df


def find_column(df, candidates):
    """候補リストからDataFrameに存在する列名を探す"""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def compare(file_a, file_b, sheet_a=0, sheet_b=0):
    df_a = load_excel(file_a, sheet_a)
    df_b = load_excel(file_b, sheet_b)

    # 列名の自動検出
    name_col_a = find_column(df_a, ["品名", "商品名", "品目", "item", "Item"])
    qty_col_a  = find_column(df_a, ["数量", "qty", "Qty", "quantity"])
    amt_col_a  = find_column(df_a, ["金額", "amount", "Amount", "価格", "単価合計"])

    name_col_b = find_column(df_b, ["品名", "商品名", "品目", "item", "Item"])
    qty_col_b  = find_column(df_b, ["数量", "qty", "Qty", "quantity"])
    amt_col_b  = find_column(df_b, ["金額", "amount", "Amount", "価格", "単価合計"])

    if not all([name_col_a, qty_col_a, amt_col_a, name_col_b, qty_col_b, amt_col_b]):
        print("【エラー】必要な列が見つかりません。")
        print(f"  ファイルA の列: {list(df_a.columns)}")
        print(f"  ファイルB の列: {list(df_b.columns)}")
        print("  「品名」「数量」「金額」に相当する列名を確認してください。")
        return

    # 必要な列だけ抽出・リネーム
    a = df_a[[name_col_a, qty_col_a, amt_col_a]].copy()
    b = df_b[[name_col_b, qty_col_b, amt_col_b]].copy()
    a.columns = ["品名", "数量", "金額"]
    b.columns = ["品名", "数量", "金額"]

    a["品名"] = a["品名"].astype(str).str.strip()
    b["品名"] = b["品名"].astype(str).str.strip()

    # 品名でマージ（outer join で抜け落ちも検出）
    merged = pd.merge(a, b, on="品名", how="outer", suffixes=("_A", "_B"))

    issues = []

    for _, row in merged.iterrows():
        name = row["品名"]
        qty_a  = row.get("数量_A")
        qty_b  = row.get("数量_B")
        amt_a  = row.get("金額_A")
        amt_b  = row.get("金額_B")

        if pd.isna(qty_a) and pd.isna(amt_a):
            issues.append({"品名": name, "種別": "【Aのみ欠落】Bにあり・Aになし",
                           "数量A": "-", "数量B": qty_b, "金額A": "-", "金額B": amt_b})
        elif pd.isna(qty_b) and pd.isna(amt_b):
            issues.append({"品名": name, "種別": "【Bのみ欠落】Aにあり・Bになし",
                           "数量A": qty_a, "数量B": "-", "金額A": amt_a, "金額B": "-"})
        else:
            diff_qty = (qty_a != qty_b) if not (pd.isna(qty_a) or pd.isna(qty_b)) else True
            diff_amt = (abs(float(amt_a) - float(amt_b)) > 0.01) if not (pd.isna(amt_a) or pd.isna(amt_b)) else True
            if diff_qty or diff_amt:
                issues.append({"品名": name, "種別": "【不一致】",
                               "数量A": qty_a, "数量B": qty_b,
                               "金額A": amt_a, "金額B": amt_b})

    # 合計金額
    total_a = a["金額"].sum()
    total_b = b["金額"].sum()

    print("=" * 60)
    print("  Excel 突き合わせ結果")
    print("=" * 60)
    print(f"  ファイルA: {Path(file_a).name}  合計金額: {total_a:,.0f} 円")
    print(f"  ファイルB: {Path(file_b).name}  合計金額: {total_b:,.0f} 円")
    print(f"  合計差額: {total_a - total_b:+,.0f} 円")
    print("=" * 60)

    if not issues:
        print("  差異なし：すべての品名・数量・金額が一致しています。")
    else:
        print(f"  差異 {len(issues)} 件\n")
        for i, issue in enumerate(issues, 1):
            print(f"  [{i}] {issue['種別']}  品名: {issue['品名']}")
            print(f"       数量  A={issue['数量A']}  /  B={issue['数量B']}")
            print(f"       金額  A={issue['金額A']}  /  B={issue['金額B']}")
            print()

    # 結果をExcelに保存
    out_path = "突き合わせ結果.xlsx"
    result_df = pd.DataFrame(issues) if issues else pd.DataFrame(
        columns=["品名", "種別", "数量A", "数量B", "金額A", "金額B"])
    summary_df = pd.DataFrame([
        {"項目": "ファイルA", "合計金額": total_a},
        {"項目": "ファイルB", "合計金額": total_b},
        {"項目": "差額",     "合計金額": total_a - total_b},
    ])
    with pd.ExcelWriter(out_path) as writer:
        result_df.to_excel(writer, sheet_name="差異一覧", index=False)
        summary_df.to_excel(writer, sheet_name="合計サマリ", index=False)
    print(f"  結果を保存しました: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("使い方: python3 excel_compare.py ファイルA.xlsx ファイルB.xlsx")
        print("        python3 excel_compare.py ファイルA.xlsx ファイルB.xlsx シートA シートB")
        sys.exit(1)

    file_a = sys.argv[1]
    file_b = sys.argv[2]
    sheet_a = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    sheet_b = int(sys.argv[4]) if len(sys.argv) > 4 else 0

    compare(file_a, file_b, sheet_a, sheet_b)
