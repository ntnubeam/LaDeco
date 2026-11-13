#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
check_compustat_merge.py

檢查「Compustat_NA_Global_2008-2025」是否為以下兩份 WRDS 匯出資料
「正確合併 + 去除重複」後的版本（不缺漏、不擅自新增、不竄改數值）：

    NA 2008-2025      comp_na_daily_all.funda   (datafmt = STD)
    Global 2008-2025  comp_global_daily.g_funda (datafmt = HIST_STD)

檢查項目
  1. 欄位    合併檔是否包含兩份來源欄位的聯集；有沒有少欄、多欄、同一欄被拆成兩欄
  2. 缺漏    來源中的每一個 (gvkey, datadate) 是否都出現在合併檔
  3. 新增    合併檔中是否有來源裡根本不存在的 (gvkey, datadate)
  4. 去重    合併檔中同一 (gvkey, datadate) 是否只剩一列；NA/Global 重疊時保留了哪一邊
  5. 數值    合併檔每一列是否與來源中對應那一列「逐欄完全相同」
  6. 格式    gvkey/cik/sic 前導零、Excel 1,048,576 列上限截斷、來源篩選條件、日期範圍

用法（Windows PowerShell，在本檔所在資料夾執行）
    python check_compustat_merge.py
    python check_compustat_merge.py --dir "D:\AI ecosystem\財務資料"
    python check_compustat_merge.py `
        --na     "D:\AI ecosystem\財務資料\NA 2008-2025.xlsx" `
        --global "D:\AI ecosystem\財務資料\Global 2008-2025.xlsx" `
        --merged "D:\AI ecosystem\財務資料\Compustat_NA_Global_2008-2025.csv"

需要：Python 3.8+、pandas、numpy、openpyxl
建議：pip install python-calamine   （讀 .xlsx 會快很多）

本程式只讀取檔案，不會修改來源檔或合併檔。
報告輸出到 <資料夾>\merge_check_report_<時間>\
"""

import argparse
import copy
import io
import os
import re
import sys
import time
import zipfile

import numpy as np
import pandas as pd

# ----------------------------------------------------------------------------
# 設定
# ----------------------------------------------------------------------------
DEFAULT_DIR = r"D:\AI ecosystem\財務資料"
DEFAULT_NA_NAME = "NA 2008-2025.xlsx"
DEFAULT_GLOBAL_NAME = "Global 2008-2025.xlsx"
MERGED_NAME_HINT = "compustat_na_global"

DATA_EXTS = (".csv", ".txt", ".tsv", ".xlsx", ".xlsm", ".xls",
             ".parquet", ".dta", ".sas7bdat", ".zip")
EXCEL_MAX_DATA_ROWS = 1048575          # Excel 上限 1,048,576 列 − 標題列
CHUNK = 200_000

DATE_MIN, DATE_MAX = "2008-01-01", "2025-12-31"
EXPECTED_FILTERS = {                    # 來自兩份 .sql 的 WHERE 條件
    "NA": {"consol": {"C"}, "indfmt": {"INDL"}, "datafmt": {"STD"},
           "curcd": {"USD", "CAD"}, "costat": {"A", "I"}},
    "Global": {"consol": {"C"}, "indfmt": {"INDL"}, "datafmt": {"HIST_STD"},
               "costat": {"A", "I"}},
}

NUMERIC_COLS = {
    "act", "at", "che", "dlc", "dltt", "lct", "ppent", "pstk", "rect", "seq",
    "cogs", "ebitda", "ni", "sale", "xad", "xrd", "xsga", "capx", "csho", "emp",
    "teq", "prcc_f", "ceq", "ib", "icapt", "invt", "lt", "nit", "oancf", "oiadp",
    "revt", "nicon",
}
DATE_COLS = {"datadate", "ipodate"}
ZFILL_COLS = {"gvkey": 6, "cik": 10, "sic": 4}      # 這些代碼有前導零
INT_COLS = {"gind", "gsubind", "naics", "exchg", "fyr", "fyear"}
MISSING_TOKENS = {"", "nan", "NaN", "NAN", "NaT", "None", "<NA>", "null", "NULL", "."}

KEY_DESC = "(gvkey, datadate)"

_T0 = time.time()


def log(msg):
    print("[%7.1fs] %s" % (time.time() - _T0, msg), flush=True)


# ----------------------------------------------------------------------------
# 欄名與數值正規化
# ----------------------------------------------------------------------------
_HDR_RE = re.compile(r"^\(([A-Za-z0-9_]+)\)")


def col_code(header):
    """'(gvkey) Global Company Key - Company' -> 'gvkey'；'GVKEY' -> 'gvkey'"""
    h = str(header).replace("\ufeff", "").strip()
    m = _HDR_RE.match(h)
    return (m.group(1) if m else h).strip().lower()


def _clean(s):
    s = s.astype(object).where(s.notna(), "")
    s = s.astype(str).str.strip()
    return s.mask(s.isin(MISSING_TOKENS), "")


_INTLIKE = r"^[+-]?\d+\.0*$"
_SCI = r"^[+-]?\d+(\.\d+)?[eE][+-]?\d+$"


def _intlike(s):
    """'1004.0' -> '1004'；'1.2E+05' -> '120000'"""
    s = s.copy()
    m = s.str.match(_INTLIKE)
    if m.any():
        s[m] = s[m].str.replace(r"\.0*$", "", regex=True)
    m = s.str.match(_SCI)
    if m.any():
        s[m] = s[m].map(lambda x: str(int(round(float(x)))))
    return s


_DATE_RULES = [
    (r"^\d{4}-\d{1,2}-\d{1,2}$", "%Y-%m-%d"),
    (r"^\d{4}/\d{1,2}/\d{1,2}$", "%Y/%m/%d"),
    (r"^\d{8}$", "%Y%m%d"),
    (r"^\d{1,2}/\d{1,2}/\d{4}$", "%m/%d/%Y"),
    (r"^\d{1,2}[A-Za-z]{3}\d{4}$", "%d%b%Y"),
]


def _dates(s):
    """各種日期寫法 -> 'YYYY-MM-DD'；無法辨識的保留原字串（之後會被當成不一致）。"""
    s = _intlike(s)
    s = s.str.replace(r"[ T]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?$", "", regex=True)
    out = s.astype(object).copy()
    for pat, fmt in _DATE_RULES:
        m = s.str.match(pat)
        if m.any():
            d = pd.to_datetime(s[m], format=fmt, errors="coerce")
            ok = d.notna().to_numpy()
            if ok.any():
                out.loc[d.index[ok]] = d[ok].dt.strftime("%Y-%m-%d").to_numpy()
    m = s.str.match(r"^\d{5}$")          # Excel 日期序號
    if m.any():
        d = pd.to_datetime(s[m].astype("int64"), unit="D", origin="1899-12-30")
        out.loc[d.index] = d.dt.strftime("%Y-%m-%d").to_numpy()
    return out


class FileStats:
    def __init__(self, label, path):
        self.label = label
        self.path = path
        self.notes = []
        self.part_rows = []                       # [(名稱, 列數)]
        self.short = {k: 0 for k in ZFILL_COLS}   # 前導零遺失筆數
        self.bad_num = {}                         # 數值欄出現非數字
        self.dup_cols = {}                        # 同一代碼出現多欄
        self.labels = {}                          # 代碼 -> 原始欄名
        self.date_samples = []


def _convert(code, s, st):
    if code in DATE_COLS:
        if code == "datadate" and len(st.date_samples) < 3:
            st.date_samples.extend([v for v in s.head(3).tolist() if v][: 3 - len(st.date_samples)])
        return _dates(s)
    if code in ZFILL_COLS:
        n = ZFILL_COLS[code]
        t = _intlike(s)
        dig = t.str.match(r"^\d+$")
        st.short[code] += int((dig & (t.str.len() < n)).sum())
        if dig.any():
            t[dig] = t[dig].str.zfill(n)
        return t.astype(object)
    if code in INT_COLS:
        return _intlike(s).astype(object)
    if code in NUMERIC_COLS:
        x = pd.to_numeric(s.str.replace(",", "", regex=False), errors="coerce").astype("float64")
        bad = int(((s != "").to_numpy() & x.isna().to_numpy()).sum())
        if bad:
            st.bad_num[code] = st.bad_num.get(code, 0) + bad
        return x
    return s.astype(object)


def normalize(raw, st):
    """原始字串表 -> 欄名換成代碼、數值/日期/代碼統一格式。"""
    raw_cols = list(raw.columns)
    codes = [col_code(c) for c in raw_cols]
    out = {}
    for code in dict.fromkeys(codes):
        idxs = [i for i, c in enumerate(codes) if c == code]
        st.labels.setdefault(code, [str(raw_cols[i]) for i in idxs])
        vals = [_convert(code, _clean(raw.iloc[:, i]), st).reset_index(drop=True) for i in idxs]
        v = vals[0]
        if len(vals) > 1:   # 例：NA 的 "(gvkey) Global Company Key" 與 Global 的 "... - Company" 沒對齊
            info = st.dup_cols.setdefault(code, {"labels": [str(raw_cols[i]) for i in idxs], "conflicts": 0})
            for o in vals[1:]:
                if v.dtype.kind == "f":
                    a, b = v.to_numpy(), o.to_numpy()
                    conf = ~np.isnan(a) & ~np.isnan(b) & (a != b)
                    v = v.fillna(o)
                else:
                    conf = ((v != "") & (o != "") & (v != o)).to_numpy()
                    v = v.where(v != "", o)
                info["conflicts"] += int(conf.sum())
        out[code] = v
    return pd.DataFrame(out)


# ----------------------------------------------------------------------------
# 讀檔（csv / xlsx / zip / 資料夾 / parquet / dta / sas7bdat）
# ----------------------------------------------------------------------------
def _sniff_sep(head):
    line = head.split(b"\n", 1)[0]
    counts = {sep: line.count(sep.encode()) for sep in (",", "\t", ";", "|")}
    return max(counts, key=counts.get)


def _read_delimited(opener, name, st):
    with opener() as f:
        head = f.read(65536)
    sep = _sniff_sep(head)
    last_err = None
    for enc in ("utf-8-sig", "cp950", "latin-1"):
        snapshot = copy.deepcopy(st.__dict__)
        try:
            parts, n = [], 0
            with opener() as fb:
                ft = io.TextIOWrapper(fb, encoding=enc, newline="")
                reader = pd.read_csv(ft, sep=sep, dtype=str, keep_default_na=False,
                                     na_filter=False, chunksize=CHUNK)
                for chunk in reader:
                    parts.append(normalize(chunk, st))
                    n += len(chunk)
                    log("  %s：已讀 %s 列" % (name, format(n, ",")))
            st.notes.append("%s：編碼 %s，分隔符號 %r" % (name, enc, sep))
            st.part_rows.append((name, n))
            return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
        except UnicodeDecodeError as e:
            st.__dict__.clear()
            st.__dict__.update(snapshot)
            last_err = e
            log("  %s：不是 %s 編碼，改用下一種重讀" % (name, enc))
    raise last_err


def _rows_to_frames(row_iter, name, sheet_title, st):
    """把逐列 iterator 切塊轉成正規化後的 DataFrame。回傳 (frames, 列數)；無 gvkey/datadate 則回 None。"""
    header = None
    for r in row_iter:
        if r is not None and any(v is not None and str(v).strip() != "" for v in r):
            header = list(r)
            break
    if header is None:
        return None
    while header and (header[-1] is None or str(header[-1]).strip() == ""):
        header.pop()
    ncol = len(header)
    hdr = [str(h) if h is not None and str(h).strip() != "" else "_unnamed_%d" % i
           for i, h in enumerate(header)]
    codes = {col_code(h) for h in hdr}
    if not {"gvkey", "datadate"} <= codes:
        st.notes.append("%s 工作表「%s」沒有 gvkey / datadate 欄，已略過" % (name, sheet_title))
        return None
    frames, buf, n = [], [], 0
    for r in row_iter:
        if r is None:
            continue
        r = list(r[:ncol])
        if len(r) < ncol:
            r += [None] * (ncol - len(r))
        if all(v is None or (isinstance(v, str) and v.strip() == "") for v in r):
            continue
        buf.append(r)
        if len(buf) >= CHUNK:
            frames.append(normalize(pd.DataFrame(buf, columns=hdr, dtype=object), st))
            n += len(buf)
            buf = []
            log("  %s「%s」：已讀 %s 列" % (name, sheet_title, format(n, ",")))
    if buf:
        frames.append(normalize(pd.DataFrame(buf, columns=hdr, dtype=object), st))
        n += len(buf)
    log("  %s「%s」：共 %s 列" % (name, sheet_title, format(n, ",")))
    return frames, n


def _read_xlsx(src, name, st):
    """逐工作表、逐塊讀取 .xlsx（有 python-calamine 用它，否則 openpyxl read-only）。"""
    frames = []
    try:
        import python_calamine
        wb = (python_calamine.CalamineWorkbook.from_path(src) if isinstance(src, str)
              else python_calamine.CalamineWorkbook.from_filelike(src))
        for sname in wb.sheet_names:
            res = _rows_to_frames(iter(wb.get_sheet_by_name(sname).iter_rows()), name, sname, st)
            if res:
                frames.extend(res[0])
                st.part_rows.append(("%s / 工作表「%s」" % (name, sname), res[1]))
        st.notes.append("%s：以 python-calamine 讀取" % name)
        return frames
    except ImportError:
        pass
    except Exception as e:  # calamine 失敗就退回 openpyxl
        log("  python-calamine 讀取失敗（%s），改用 openpyxl" % e)
        frames = []
        st.part_rows = [p for p in st.part_rows if not p[0].startswith(name + " / ")]
        if not isinstance(src, str):
            src.seek(0)
    import openpyxl
    log("  %s：使用 openpyxl 讀取（大檔可能要數分鐘；pip install python-calamine 可加速）" % name)
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            res = _rows_to_frames(ws.iter_rows(values_only=True), name, ws.title, st)
            if res:
                frames.extend(res[0])
                st.part_rows.append(("%s / 工作表「%s」" % (name, ws.title), res[1]))
    finally:
        wb.close()
    st.notes.append("%s：以 openpyxl 讀取" % name)
    return frames


def _frame_to_str(df):
    out = {}
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            out[c] = s.dt.strftime("%Y-%m-%d").astype(object).where(s.notna(), "")
        elif pd.api.types.is_float_dtype(s):
            out[c] = s.map(lambda v: "" if pd.isna(v) else repr(float(v)))
        else:
            out[c] = s.astype(object).where(s.notna(), "").astype(str)
    return pd.DataFrame(out)


def _read_one(name, ext, opener, path, st):
    if ext in (".csv", ".txt", ".tsv"):
        return [_read_delimited(opener, name, st)]
    if ext in (".xlsx", ".xlsm"):
        if path:
            return _read_xlsx(path, name, st)
        with opener() as f:
            return _read_xlsx(io.BytesIO(f.read()), name, st)
    if ext in (".xls", ".parquet", ".dta", ".sas7bdat"):
        src = path
        if not path:
            with opener() as f:
                src = io.BytesIO(f.read())
        if ext == ".xls":
            sheets = pd.read_excel(src, sheet_name=None, dtype=str, keep_default_na=False, na_filter=False)
        elif ext == ".parquet":
            sheets = {"": pd.read_parquet(src)}
        elif ext == ".dta":
            sheets = {"": pd.read_stata(src)}
        else:
            sheets = {"": pd.read_sas(src, format="sas7bdat", encoding="latin-1")}
        frames = []
        for sname, df in sheets.items():
            if not {"gvkey", "datadate"} <= {col_code(c) for c in df.columns}:
                st.notes.append("%s 工作表「%s」沒有 gvkey / datadate 欄，已略過" % (name, sname))
                continue
            df = _frame_to_str(df)
            for i in range(0, len(df), CHUNK):
                frames.append(normalize(df.iloc[i:i + CHUNK].reset_index(drop=True), st))
            st.part_rows.append(("%s%s" % (name, (" / 工作表「%s」" % sname) if sname else ""), len(df)))
        return frames
    raise ValueError("不支援的檔案格式：%s" % name)


def read_dataset(path, label):
    st = FileStats(label, path)
    log("讀取 %s：%s" % (label, path))
    frames = []
    if os.path.isdir(path):
        files = []
        for root, _, fs in os.walk(path):
            for f in fs:
                if f.lower().endswith(DATA_EXTS) and not f.startswith("~$"):
                    files.append(os.path.join(root, f))
        if not files:
            raise SystemExit("[錯誤] 資料夾內找不到資料檔：%s" % path)
        if len(files) > 1:
            st.notes.append("資料夾內有 %d 個檔案，全部合併讀取：%s"
                            % (len(files), "; ".join(os.path.relpath(f, path) for f in sorted(files))))
        for f in sorted(files):
            frames += _read_path(f, os.path.relpath(f, path), st)
    else:
        frames += _read_path(path, os.path.basename(path), st)
    frames = [f for f in frames if len(f.columns)]
    if not frames:
        raise SystemExit("[錯誤] %s 讀不到任何含 gvkey / datadate 的資料" % label)
    df = pd.concat(frames, ignore_index=True)
    for c in df.columns:     # 不同分段欄位不一致時補空值
        if df[c].dtype == object:
            df[c] = df[c].where(df[c].notna(), "")
    for need in ("gvkey", "datadate"):
        if need not in df.columns:
            raise SystemExit("[錯誤] %s 找不到 %s 欄位（欄名：%s）" % (label, need, list(df.columns)[:20]))
    log("  %s：共 %s 列、%d 欄" % (label, format(len(df), ","), df.shape[1]))
    return df, st


def _read_path(path, name, st):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".zip":
        frames = []
        with zipfile.ZipFile(path) as zf:
            members = [m for m in zf.namelist()
                       if m.lower().endswith(DATA_EXTS) and not m.endswith("/")
                       and not m.startswith("__MACOSX") and not os.path.basename(m).startswith("~$")]
            if not members:
                raise SystemExit("[錯誤] zip 內沒有資料檔：%s" % path)
            if len(members) > 1:
                st.notes.append("%s 內有 %d 個資料檔，全部合併讀取：%s" % (name, len(members), "; ".join(members)))
            for m in members:
                mext = os.path.splitext(m)[1].lower()
                frames += _read_one("%s/%s" % (name, m), mext, (lambda m=m: zf.open(m)), None, st)
        return frames
    return _read_one(name, ext, (lambda: open(path, "rb")), path, st)


# ----------------------------------------------------------------------------
# 比對工具
# ----------------------------------------------------------------------------
def make_key(df):
    g = df["gvkey"].astype(object)
    if "conm" in df.columns:
        g = g.where(g != "", "NOGVKEY:" + df["conm"].astype(object).astype(str))
    return (g + "|" + df["datadate"].astype(object)).to_numpy(dtype=object)


def canon(df, c):
    """單一欄位的標準字串表示（數值取 12 位有效數字，消除浮點雜訊）。"""
    if c not in df.columns:
        return np.full(len(df), "", dtype=object)
    s = df[c]
    if s.dtype.kind == "f":
        v = s.to_numpy(dtype="float64")
        out = np.char.mod("%.12g", v).astype(object)
        out[np.isnan(v)] = ""
        out[out == "-0"] = "0"
        return out
    out = s.to_numpy(dtype=object)
    nulls = pd.isna(out)
    if nulls.any():
        out = out.copy()
        out[nulls] = ""
    return out


def row_hash(df, cols):
    h = np.zeros(len(df), dtype=np.uint64)
    with np.errstate(over="ignore"):
        for i, c in enumerate(cols):
            hc = pd.util.hash_array(canon(df, c))
            h = h * np.uint64(1099511628211) + hc + np.uint64(i + 1)
    return h


def isin(values, pool):
    return pd.Index(values).isin(pool) if len(values) else np.zeros(0, dtype=bool)


def fmt(n):
    return format(int(n), ",")


def fmt_vc(vc):
    return "、".join("%s %s" % (k if k != "" else "(空白)", fmt(v)) for k, v in vc.items())


class Report:
    def __init__(self):
        self.lines, self.errors, self.warnings = [], [], []

    def h(self, title):
        self.lines += ["", "=" * 78, title, "=" * 78]

    def p(self, text=""):
        self.lines.append(text)

    def err(self, text):
        self.errors.append(text)
        self.p("  [錯誤] " + text)

    def warn(self, text):
        self.warnings.append(text)
        self.p("  [警告] " + text)

    def ok(self, text):
        self.p("  [OK]   " + text)


def save_csv(df, outdir, name):
    path = os.path.join(outdir, name)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return name


# ----------------------------------------------------------------------------
# 主要檢查
# ----------------------------------------------------------------------------
def source_health(name, df, k, st, R, is_source=True):
    R.p()
    R.p("【%s】 %s 列、%d 欄" % (name, fmt(len(df)), df.shape[1]))
    for part, n in st.part_rows:
        R.p("    - %s：%s 列" % (part, fmt(n)))
        if n == EXCEL_MAX_DATA_ROWS:
            if is_source:
                R.warn("%s 剛好 %s 列 = Excel 單一工作表上限，來源檔本身很可能已被截斷。"
                       "就算合併檔與它完全一致，也會少掉 WRDS 上超出上限的資料；"
                       "建議從 WRDS 改用 CSV 格式重新下載後再合併" % (part, fmt(n)))
            else:
                R.err("%s 剛好 %s 列 = Excel 單一工作表上限，合併檔很可能被截斷（後面的列不見了）"
                      % (part, fmt(n)))
    for note in st.notes:
        R.p("    · " + note)

    dd = df["datadate"]
    valid = dd.str.match(r"^\d{4}-\d{2}-\d{2}$").to_numpy()
    if valid.any():
        R.p("    datadate 範圍：%s ~ %s" % (dd[valid].min(), dd[valid].max()))
    bad_date = int((~valid).sum())
    if bad_date:
        R.warn("%s：%s 列 datadate 空白或無法辨識" % (name, fmt(bad_date)))
    oor = int((valid & ((dd < DATE_MIN) | (dd > DATE_MAX)).to_numpy()).sum())
    if oor:
        R.warn("%s：%s 列 datadate 不在 %s ~ %s" % (name, fmt(oor), DATE_MIN, DATE_MAX))

    filters = EXPECTED_FILTERS.get(name)
    if filters is None:   # 合併檔：允許兩邊的值
        filters = {c: EXPECTED_FILTERS["NA"].get(c, set()) | EXPECTED_FILTERS["Global"].get(c, set())
                   for c in ("consol", "indfmt", "datafmt", "costat")}
    for col, allowed in filters.items():
        if col not in df.columns:
            continue
        vc = df[col].value_counts()
        bad = vc[~vc.index.isin(list(allowed))]
        if len(bad):
            R.warn("%s：%s 出現預期外的值 %s（預期 %s）" % (name, col, fmt_vc(bad.head(10)), sorted(allowed)))
    if "datafmt" in df.columns:
        R.p("    datafmt 分布：%s" % fmt_vc(df["datafmt"].value_counts()))

    empty_g = int((df["gvkey"] == "").sum())
    if empty_g:
        R.warn("%s：%s 列 gvkey 空白" % (name, fmt(empty_g)))

    dupk = pd.Series(k).duplicated(keep=False)
    if is_source and dupk.any():   # 合併檔的重複在第 3 節檢查
        R.warn("%s：同一 %s 出現多列，涉及 %s 列 / %s 組" % (
            name, KEY_DESC, fmt(dupk.sum()), fmt(pd.Series(k[dupk.to_numpy()]).nunique())))
    exact_dup = int(pd.Series(row_hash(df, list(df.columns))).duplicated().sum()) if is_source else 0
    if exact_dup:
        R.warn("%s：有 %s 列是「整列完全重複」" % (name, fmt(exact_dup)))
    if "fyear" in df.columns and is_source:
        fy = df["fyear"] != ""
        k2 = (df["gvkey"] + "|" + df["fyear"])[fy]
        n2 = int(k2.duplicated(keep=False).sum())
        if n2:
            R.p("    註：%s 列的 (gvkey, fyear) 重複（會計年度結束月變更所致，屬正常資料）。"
                "若合併時用 gvkey+fyear 去重，這些列會被誤刪。" % fmt(n2))

    for c, n in st.short.items():
        if n:
            msg = "%s：%s 筆 %s 少了前導零（例如 001004 變成 1004），常見原因是用 Excel 開啟後另存" % (
                name, fmt(n), c)
            (R.p if is_source else R.warn)(("    註：" + msg) if is_source else msg)
    for c, n in st.bad_num.items():
        R.warn("%s：數值欄 %s 有 %s 筆不是數字" % (name, c, fmt(n)))
    for c, info in st.dup_cols.items():
        msg = "%s：%s 對應到 %d 個欄位 %s（已合併成一欄比對），其中 %s 列兩欄的值互相矛盾" % (
            name, c, len(info["labels"]), info["labels"], fmt(info["conflicts"]))
        if info["conflicts"]:
            R.err(msg)
        else:
            R.warn(msg + "。合併時欄名沒對齊，造成同一變數被拆成多欄")


def analyze(NA, G, M, stats, outdir, max_examples):
    R = Report()
    files_out = []

    # ---------------- 1. 欄位 ----------------
    R.h("1. 欄位檢查")
    union = list(NA.columns) + [c for c in G.columns if c not in NA.columns]
    missing_cols = [c for c in union if c not in M.columns]
    extra_cols = [c for c in M.columns if c not in union]
    C = [c for c in union if c in M.columns]
    R.p("  NA %d 欄、Global %d 欄、聯集 %d 欄（共同 %d、NA 獨有 %d、Global 獨有 %d）；合併檔 %d 欄" % (
        NA.shape[1], G.shape[1], len(union), len(set(NA.columns) & set(G.columns)),
        len(set(NA.columns) - set(G.columns)), len(set(G.columns) - set(NA.columns)), M.shape[1]))
    R.p("  NA 獨有欄位：%s" % sorted(set(NA.columns) - set(G.columns)))
    R.p("  Global 獨有欄位：%s" % sorted(set(G.columns) - set(NA.columns)))
    if missing_cols:
        R.err("合併檔缺少欄位：%s" % missing_cols)
    else:
        R.ok("合併檔包含兩份來源的所有欄位")
    if extra_cols:
        R.warn("合併檔多出來源沒有的欄位：%s（若是自己加的來源標記欄可忽略；這些欄不納入比對）" % extra_cols)

    # ---------------- 2. 來源與合併檔健檢 ----------------
    R.h("2. 檔案健檢")
    kNA, kG, kM = make_key(NA), make_key(G), make_key(M)
    source_health("NA", NA, kNA, stats["NA"], R)
    source_health("Global", G, kG, stats["Global"], R)
    source_health("合併檔", M, kM, stats["M"], R, is_source=False)

    # ---------------- 3. 列與鍵值核對 ----------------
    R.h("3. 合併結果核對（以 %s 為鍵）" % KEY_DESC)
    log("計算每列指紋（逐欄比對用）...")
    hNA, hG, hM = row_hash(NA, C), row_hash(G, C), row_hash(M, C)
    M_inNA, M_inG = np.isin(hM, hNA), np.isin(hM, hG)
    M_exact = M_inNA | M_inG
    NA_inM, G_inM = np.isin(hNA, hM), np.isin(hG, hM)

    iNA, iG, iM = pd.Index(pd.unique(kNA)), pd.Index(pd.unique(kG)), pd.Index(pd.unique(kM))
    overlap = iNA.intersection(iG)
    allk = iNA.union(iG)
    missing = allk.difference(iM)
    extra = iM.difference(allk)

    R.p("  NA 列數                         %12s" % fmt(len(NA)))
    R.p("  Global 列數                     %12s" % fmt(len(G)))
    R.p("  NA 與 Global 重疊的 %s %10s" % (KEY_DESC, fmt(len(overlap))))
    R.p("  預期合併去重後列數（不重複鍵）  %12s" % fmt(len(allk)))
    R.p("  合併檔實際列數                  %12s" % fmt(len(M)))
    R.p("  合併檔不重複鍵數                %12s" % fmt(len(iM)))
    R.p()

    # 缺漏
    if len(missing):
        m_na = missing.isin(iNA) & ~missing.isin(iG)
        m_g = missing.isin(iG) & ~missing.isin(iNA)
        m_both = missing.isin(overlap)
        R.err("缺漏：%s 組 %s 在來源有、合併檔沒有（只在 NA %s、只在 Global %s、兩邊都有 %s）" % (
            fmt(len(missing)), KEY_DESC, fmt(m_na.sum()), fmt(m_g.sum()), fmt(m_both.sum())))
        rows = pd.concat([
            NA.loc[isin(kNA, missing)].assign(_來源="NA"),
            G.loc[isin(kG, missing)].assign(_來源="Global"),
        ], ignore_index=True)
        files_out.append(save_csv(rows, outdir, "1_缺漏_來源有但合併檔沒有.csv"))
    else:
        R.ok("沒有缺漏：來源的每一個 %s 都在合併檔中" % KEY_DESC)

    # 擅自新增
    if len(extra):
        R.err("擅自新增：合併檔有 %s 組 %s 在兩份來源中都找不到" % (fmt(len(extra)), KEY_DESC))
        files_out.append(save_csv(M.loc[isin(kM, extra)], outdir, "2_新增_來源沒有的列.csv"))
    else:
        R.ok("沒有擅自新增的列：合併檔每一個 %s 都來自 NA 或 Global" % KEY_DESC)

    # 去重
    exact_dup_rows = int(pd.Series(hM).duplicated().sum())
    dupk_mask = pd.Series(kM).duplicated(keep=False).to_numpy()
    n_dup_keys = int(pd.Series(kM[dupk_mask]).nunique()) if dupk_mask.any() else 0
    if exact_dup_rows:
        R.err("未去重：合併檔有 %s 列是其他列的「整列完全重複」" % fmt(exact_dup_rows))
    if n_dup_keys:
        R.err("未去重：合併檔有 %s 組 %s 出現不只一次（共 %s 列）" % (
            fmt(n_dup_keys), KEY_DESC, fmt(dupk_mask.sum())))
        lab = np.where(M_inNA, "與NA某列完全相同", np.where(M_inG, "與Global某列完全相同", "與來源不完全相同"))
        d = M.loc[dupk_mask].assign(_比對結果=lab[dupk_mask])
        d = d.sort_values(["gvkey", "datadate"], kind="mergesort")
        files_out.append(save_csv(d, outdir, "3_未去重_同鍵多列.csv"))
    if not exact_dup_rows and not n_dup_keys:
        R.ok("已去重：合併檔每一個 %s 只出現一次" % KEY_DESC)

    # 重疊鍵的保留方式
    keys_exact_NA = pd.Index(pd.unique(kM[M_inNA]))
    keys_exact_G = pd.Index(pd.unique(kM[M_inG]))
    keys_exact = keys_exact_NA.union(keys_exact_G)
    R.p()
    R.p("  NA 與 Global 重疊的 %s 組，合併檔保留方式：" % fmt(len(overlap)))
    if len(overlap):
        kn = overlap.isin(keys_exact_NA)
        kg = overlap.isin(keys_exact_G)
        inm = overlap.isin(iM)
        kept = np.select([kn & kg, kn, kg, inm], ["兩邊都保留", "保留NA", "保留Global", "混合或被修改"], "整組缺漏")
        vc = pd.Series(kept).value_counts()
        for name in ("保留NA", "保留Global", "兩邊都保留", "混合或被修改", "整組缺漏"):
            if name in vc:
                R.p("    %-10s %10s" % (name, fmt(vc[name])))
        if "兩邊都保留" in vc:
            R.err("%s 組重疊鍵在合併檔中 NA 與 Global 兩列都留著（跨資料庫重複沒有去除）" % fmt(vc["兩邊都保留"]))
        if vc.get("保留NA", 0) and vc.get("保留Global", 0):
            R.warn("重疊鍵有的保留 NA、有的保留 Global（去重規則不一致；若是刻意依條件挑選可忽略）")
        NAu = NA.assign(_k=kNA).drop_duplicates("_k").set_index("_k")
        Gu = G.assign(_k=kG).drop_duplicates("_k").set_index("_k")
        a, b = NAu.reindex(overlap), Gu.reindex(overlap)
        ov = pd.DataFrame({
            "gvkey": a["gvkey"].to_numpy(), "datadate": a["datadate"].to_numpy(),
            "conm_NA": canon(a, "conm"), "conm_Global": canon(b, "conm"),
            "curcd_NA": canon(a, "curcd"), "curcd_Global": canon(b, "curcd"),
            "fic_NA": canon(a, "fic"), "fic_Global": canon(b, "fic"),
            "at_NA": canon(a, "at"), "at_Global": canon(b, "at"),
            "合併檔保留": kept,
        })
        files_out.append(save_csv(ov, outdir, "4_NA與Global重疊鍵.csv"))

    # 來源列沒有原封不動出現在合併檔的原因
    R.p()
    for name, k, inM in (("NA", kNA, NA_inM), ("Global", kG, G_inM)):
        absent = ~inM
        miss = absent & ~isin(k, iM)
        dedup = absent & ~miss & isin(k, keys_exact)
        modified = absent & ~miss & ~dedup
        R.p("  %s 的 %s 列：原封不動在合併檔 %s；因同鍵已保留另一列而被去重 %s；整鍵缺漏 %s；值被改動 %s" % (
            name, fmt(len(k)), fmt(inM.sum()), fmt(dedup.sum()), fmt(miss.sum()), fmt(modified.sum())))

    # ---------------- 4. 數值逐欄比對 ----------------
    R.h("4. 數值逐欄比對")
    R.p("  合併檔 %s 列中：與 NA 某列完全相同 %s、與 Global 某列完全相同 %s、與任何來源列都不完全相同 %s" % (
        fmt(len(M)), fmt(M_inNA.sum()), fmt(M_inG.sum()), fmt((~M_exact).sum())))
    diag = ~M_exact & isin(kM, allk)
    n_diag = int(diag.sum())
    if n_diag == 0:
        R.ok("合併檔每一列都能在 NA 或 Global 找到逐欄完全相同的列（數值未被改動）")
    else:
        log("診斷 %s 列不一致的資料..." % fmt(n_diag))
        Md = M.loc[diag].reset_index(drop=True)
        kd = kM[diag]
        NAu = NA.assign(_k=kNA).drop_duplicates("_k").set_index("_k")
        Gu = G.assign(_k=kG).drop_duplicates("_k").set_index("_k")
        has_na = np.asarray(isin(kd, NAu.index))
        has_g = np.asarray(isin(kd, Gu.index))
        NAd, Gd = NAu.reindex(kd), Gu.reindex(kd)
        nd_na = np.zeros(n_diag, dtype=np.int32)
        nd_g = np.zeros(n_diag, dtype=np.int32)
        nunexp = np.zeros(n_diag, dtype=np.int32)
        dna, dg = {}, {}
        for c in C:
            m, a, b = canon(Md, c), canon(NAd, c), canon(Gd, c)
            dna[c] = has_na & (m != a)
            dg[c] = has_g & (m != b)
            nd_na += dna[c]
            nd_g += dg[c]
            nunexp += (m != "") & ~((has_na & (m == a)) | (has_g & (m == b)))
        best_na = has_na & (~has_g | (nd_na <= nd_g))
        mixed = has_na & has_g & (nunexp == 0)
        n_mixed, n_mod = int(mixed.sum()), int((~mixed).sum())
        if n_mod:
            R.err("%s 列的數值與來源對應列不同（被修改、補值或格式轉換造成）" % fmt(n_mod))
        if n_mixed:
            R.warn("%s 列是 NA 與 Global 兩列的「混合」（每個值都來自來源，但同一列混用兩邊的值，"
                   "幣別可能不一致）" % fmt(n_mixed))
        R.p()
        R.p("  各欄不一致的列數（與最接近的來源列相比）：")
        R.p("    %-10s %12s %12s" % ("欄位", "被修改列", "混合列"))
        detail = []
        for c in C:
            dbest = np.where(best_na, dna[c], dg[c])
            n1, n2 = int((dbest & ~mixed).sum()), int((dbest & mixed).sum())
            if n1 or n2:
                R.p("    %-10s %12s %12s" % (c, fmt(n1), fmt(n2)))
            idx = np.nonzero(dbest)[0][:max_examples]
            if len(idx):
                sub = Md.iloc[idx]
                detail.append(pd.DataFrame({
                    "gvkey": sub["gvkey"].to_numpy(), "datadate": sub["datadate"].to_numpy(),
                    "conm": canon(sub, "conm"), "欄位": c,
                    "合併檔的值": canon(sub, c),
                    "對照來源": np.where(best_na[idx], "NA", "Global"),
                    "NA的值": np.where(has_na[idx], canon(NAd.iloc[idx], c), "(無此鍵)"),
                    "Global的值": np.where(has_g[idx], canon(Gd.iloc[idx], c), "(無此鍵)"),
                    "類型": np.where(mixed[idx], "混合", "被修改"),
                }))
        if detail:
            files_out.append(save_csv(pd.concat(detail, ignore_index=True), outdir,
                                      "5_數值不一致明細.csv"))
            R.p("  （每欄最多列出 %d 筆範例於 5_數值不一致明細.csv）" % max_examples)

    # ---------------- 5. 年度對照 ----------------
    R.h("5. 各年度列數對照（依 datadate 年份）")

    def by_year(keys):
        s = pd.Series(np.asarray(keys, dtype=object)).str.rsplit("|", n=1).str[-1].str[:4]
        return s.value_counts()

    yt = pd.DataFrame({
        "NA": by_year(kNA), "Global": by_year(kG), "重疊": by_year(overlap),
        "預期(去重後)": by_year(allk), "合併檔": by_year(kM),
    }).fillna(0).astype("int64").sort_index()
    yt["差異(合併檔-預期)"] = yt["合併檔"] - yt["預期(去重後)"]
    yt.index.name = "年度"
    R.p(yt.to_string())
    files_out.append(save_csv(yt.reset_index(), outdir, "6_年度對照.csv"))
    return R, files_out


# ----------------------------------------------------------------------------
# 找檔案
# ----------------------------------------------------------------------------
def _candidates(folder, pred):
    out = []
    for f in sorted(os.listdir(folder)):
        p = os.path.join(folder, f)
        if f.startswith("~$") or f.lower().startswith("merge_check_report"):
            continue
        if (os.path.isfile(p) and f.lower().endswith(DATA_EXTS)) or os.path.isdir(p):
            if pred(f.lower()):
                out.append(p)
    return out


def _pick(label, explicit, folder, default_name, pred, flag):
    if explicit:
        if not os.path.exists(explicit):
            raise SystemExit("[錯誤] 找不到 %s：%s" % (label, explicit))
        return explicit
    if default_name and os.path.isfile(os.path.join(folder, default_name)):
        return os.path.join(folder, default_name)
    c = _candidates(folder, pred)
    if len(c) == 1:
        return c[0]
    if not c:
        raise SystemExit("[錯誤] 在 %s 找不到 %s，請用 %s \"完整路徑\" 指定" % (folder, label, flag))
    raise SystemExit("[錯誤] %s 有多個候選檔案，請用 %s \"完整路徑\" 指定其中一個：\n  %s"
                     % (label, flag, "\n  ".join(c)))


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="檢查 Compustat NA + Global 合併檔是否正確（不缺漏、不新增、不竄改）")
    ap.add_argument("--dir", default=DEFAULT_DIR, help="資料夾（預設：%(default)s）")
    ap.add_argument("--na", help="NA 2008-2025 檔案（.xlsx/.csv/.zip/資料夾）")
    ap.add_argument("--global", dest="glob", help="Global 2008-2025 檔案")
    ap.add_argument("--merged", help="Compustat_NA_Global_2008-2025 合併檔")
    ap.add_argument("--out", help="報告輸出資料夾（預設在 --dir 底下）")
    ap.add_argument("--max-examples", type=int, default=200, help="每個欄位最多輸出幾筆不一致範例")
    args = ap.parse_args()

    folder = args.dir
    if not (args.na and args.glob and args.merged) and not os.path.isdir(folder):
        raise SystemExit("[錯誤] 找不到資料夾 %s；請用 --dir 或 --na/--global/--merged 指定" % folder)

    def is_merged(f):
        return MERGED_NAME_HINT in f.replace(" ", "_") or ("global" in f and re.search(r"(?<![a-z])na(?![a-z])", f))

    na_path = _pick("NA 檔", args.na, folder, DEFAULT_NA_NAME,
                    lambda f: re.search(r"(?<![a-z])na(?![a-z])", f) and not is_merged(f), "--na")
    g_path = _pick("Global 檔", args.glob, folder, DEFAULT_GLOBAL_NAME,
                   lambda f: "global" in f and not is_merged(f), "--global")
    m_path = _pick("合併檔", args.merged, folder, None, is_merged, "--merged")
    if len({os.path.abspath(na_path), os.path.abspath(g_path), os.path.abspath(m_path)}) < 3:
        raise SystemExit("[錯誤] NA / Global / 合併檔 指到了同一個檔案，請分別指定")

    print("NA     ：%s" % na_path)
    print("Global ：%s" % g_path)
    print("合併檔 ：%s" % m_path)

    NA, st_na = read_dataset(na_path, "NA")
    G, st_g = read_dataset(g_path, "Global")
    M, st_m = read_dataset(m_path, "合併檔")

    outdir = args.out or os.path.join(os.path.dirname(os.path.abspath(m_path)),
                                      "merge_check_report_" + time.strftime("%Y%m%d_%H%M%S"))
    os.makedirs(outdir, exist_ok=True)

    R, files_out = analyze(NA, G, M, {"NA": st_na, "Global": st_g, "M": st_m}, outdir, args.max_examples)

    head = ["Compustat NA + Global 合併檢查報告",
            "產生時間：%s" % time.strftime("%Y-%m-%d %H:%M:%S"),
            "NA     ：%s" % na_path, "Global ：%s" % g_path, "合併檔 ：%s" % m_path, ""]
    if R.errors:
        verdict = "結論：[不通過] 合併檔不是正確版本，發現 %d 類問題：" % len(R.errors)
    elif R.warnings:
        verdict = "結論：[通過] 內容正確（無缺漏、無新增、無竄改、已去重），另有 %d 項警告請確認：" % len(R.warnings)
    else:
        verdict = "結論：[通過] 合併檔正確：無缺漏、無新增、無竄改、已去重。"
    summary = [verdict] + ["  - " + e for e in R.errors] + (
        ["", "警告："] + ["  - " + w for w in R.warnings] if R.warnings else [])
    tail = ["", "明細檔（%s）：" % outdir] + ["  - " + f for f in files_out]

    text = "\n".join(head + summary + R.lines + ["", "=" * 78] + summary + tail) + "\n"
    with open(os.path.join(outdir, "報告.txt"), "w", encoding="utf-8-sig") as f:
        f.write(text)
    print()
    print(text)
    print("報告已存到：%s" % os.path.join(outdir, "報告.txt"))
    return 1 if R.errors else 0


if __name__ == "__main__":
    sys.exit(main())
