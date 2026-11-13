#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
merge_compustat.py

把「NA 2008-2025」與「Global 2008-2025」合併，以 (gvkey, datadate) 去除重複，
輸出一個新的合併檔（不會覆蓋任何現有檔案）。

規則
  - 欄位：兩份來源欄位的聯集（53 欄），gvkey 統一成一欄；最後另加 source / in_both / firm_in_both
  - 去重鍵：(gvkey, datadate)。同一公司、同一資料日期在 NA 與 Global 都有時，預設保留 NA 那一列
    （可用 --prefer global 改成保留 Global）
  - 不用 (gvkey, fyear) 去重：變更會計年度結束月的公司，同一個 fyear 會有兩個 datadate，兩列都是真實資料
  - 數值與文字一律照抄來源，不補值、不換算；gvkey / cik / sic 保留前導零
  - 只刪除「整列完全相同」的重複列

新增欄位
  source        此列來自 NA 或 Global
  in_both       此 (gvkey, datadate) 在 NA 與 Global 都有（1 / 0）
  firm_in_both  此 gvkey 在 NA 與 Global 都出現過（1 / 0）

用法（與 check_compustat_merge.py 放在同一個資料夾）
  py merge_compustat.py
  py merge_compustat.py --out "D:\AI ecosystem\財務資料\Compustat_NA_Global_2008-2025_v2.csv"
"""

import argparse
import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import check_compustat_merge as chk  # noqa: E402  共用讀檔與正規化


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="合併 Compustat NA 與 Global，以 (gvkey, datadate) 去重")
    ap.add_argument("--dir", default=chk.DEFAULT_DIR, help="資料夾（預設：%(default)s）")
    ap.add_argument("--na", help="NA 檔（預設：<dir>\\%s）" % chk.DEFAULT_NA_NAME)
    ap.add_argument("--global", dest="glob", help="Global 檔（預設：<dir>\\%s）" % chk.DEFAULT_GLOBAL_NAME)
    ap.add_argument("--out", help="輸出檔（預設：<dir>\\Compustat_NA_Global_2008-2025_v2.csv）")
    ap.add_argument("--prefer", choices=["na", "global"], default="na",
                    help="重疊時保留哪一邊（預設 na）")
    args = ap.parse_args()

    na_path = args.na or os.path.join(args.dir, chk.DEFAULT_NA_NAME)
    g_path = args.glob or os.path.join(args.dir, chk.DEFAULT_GLOBAL_NAME)
    out_path = args.out or os.path.join(args.dir, "Compustat_NA_Global_2008-2025_v2.csv")
    for p in (na_path, g_path):
        if not os.path.exists(p):
            raise SystemExit("[錯誤] 找不到 %s" % p)
    if os.path.exists(out_path):
        raise SystemExit("[錯誤] 輸出檔已存在，為避免覆蓋請先改名或用 --out 指定新檔名：%s" % out_path)

    NA, st_na = chk.read_dataset(na_path, "NA")
    G, st_g = chk.read_dataset(g_path, "Global")

    # 原始欄名：NA 優先（gvkey 統一用 NA 的 "(gvkey) Global Company Key"）
    labels = {}
    for st in (st_g, st_na):
        for code, labs in st.labels.items():
            labels[code] = labs[0]

    # 只刪「整列完全相同」的重複列
    for name, df in (("NA", NA), ("Global", G)):
        dup = pd.Series(chk.row_hash(df, list(df.columns))).duplicated().to_numpy()
        if dup.any():
            chk.log("%s：刪除 %s 列整列完全重複的列" % (name, chk.fmt(dup.sum())))
            df.drop(index=df.index[dup], inplace=True)
            df.reset_index(drop=True, inplace=True)

    kNA, kG = chk.make_key(NA), chk.make_key(G)
    for name, k in (("NA", kNA), ("Global", kG)):
        d = pd.Series(k).duplicated(keep=False)
        if d.any():
            chk.log("[警告] %s 內有 %s 列同一 (gvkey, datadate) 但內容不同，全部保留，請自行確認"
                    % (name, chk.fmt(d.sum())))

    overlap = pd.Index(pd.unique(kNA)).intersection(pd.Index(pd.unique(kG)))
    if args.prefer == "na":
        first, second, k_first, k_second = NA.assign(source="NA"), G.assign(source="Global"), kNA, kG
    else:
        first, second, k_first, k_second = G.assign(source="Global"), NA.assign(source="NA"), kG, kNA
    keep_second = ~pd.Index(k_second).isin(pd.unique(k_first))
    merged = pd.concat([first, second.loc[keep_second]], ignore_index=True)

    union = list(NA.columns) + [c for c in G.columns if c not in NA.columns]
    k_out = chk.make_key(merged)
    firms_both = (set(NA["gvkey"]) & set(G["gvkey"])) - {""}
    merged["in_both"] = pd.Index(k_out).isin(overlap).astype(int)
    merged["firm_in_both"] = merged["gvkey"].isin(firms_both).astype(int)
    merged = merged[union + ["source", "in_both", "firm_in_both"]]
    for c in merged.columns:            # 一邊沒有的欄位留空白
        if merged[c].dtype == object:
            merged[c] = merged[c].where(merged[c].notna(), "")
    merged = merged.sort_values(["gvkey", "datadate"], kind="mergesort").reset_index(drop=True)

    expected = len(pd.Index(pd.unique(kNA)).union(pd.Index(pd.unique(kG))))
    chk.log("寫出 %s ..." % out_path)
    merged.rename(columns=labels).to_csv(out_path, index=False, encoding="utf-8-sig")

    print()
    print("NA 列數                     %12s" % chk.fmt(len(NA)))
    print("Global 列數                 %12s" % chk.fmt(len(G)))
    print("重疊 (gvkey, datadate)      %12s   → 保留 %s" % (chk.fmt(len(overlap)),
                                                         "NA" if args.prefer == "na" else "Global"))
    print("輸出列數                    %12s   （預期 %s）" % (chk.fmt(len(merged)), chk.fmt(expected)))
    print("輸出欄數                    %12d   （53 欄資料 + source / in_both / firm_in_both）" % merged.shape[1])
    print()
    print("已存到：%s" % out_path)
    print("下一步請用檢查程式驗證：")
    print('  py check_compustat_merge.py --merged "%s"' % out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
