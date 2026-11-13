# Compustat NA + Global 合併檢查

檢查 `Compustat_NA_Global_2008-2025` 是否為 `NA 2008-2025` 與 `Global 2008-2025`
正確合併、去除重複後的版本：**不缺漏、不擅自新增、不竄改數值**。程式只讀檔，不會修改任何檔案。

## 執行

```powershell
pip install pandas openpyxl python-calamine   # python-calamine 可選，讀 xlsx 快約 3 倍
cd "存放 check_compustat_merge.py 的資料夾"
python check_compustat_merge.py
```

預設讀取：

| 角色 | 預設路徑 |
|---|---|
| NA | `D:\AI ecosystem\財務資料\NA 2008-2025.xlsx` |
| Global | `D:\AI ecosystem\財務資料\Global 2008-2025.xlsx` |
| 合併檔 | 同資料夾中檔名含 `Compustat_NA_Global` 的檔案（.csv / .xlsx / .zip 皆可） |

路徑不同時可以用參數指定：

```powershell
python check_compustat_merge.py --na "D:\...\NA 2008-2025.xlsx" --global "D:\...\Global 2008-2025.xlsx" --merged "D:\...\Compustat_NA_Global_2008-2025.csv"
```

約 100 萬列資料需要 5–15 分鐘，記憶體約 2–3 GB。

## 檢查內容

1. **欄位**：合併檔應有兩份來源欄位的聯集，共 53 欄，其中共同 37 欄、NA 獨有 5 欄、Global 獨有 11 欄。
   程式也會檢查同一變數是否被拆成兩欄。NA 的欄名是 `(gvkey) Global Company Key`，Global 的是
   `(gvkey) Global Company Key - Company`；直接依欄名串接時，gvkey 會變成兩欄。
2. **缺漏**：來源的每個 `(gvkey, datadate)` 都必須出現在合併檔中。
3. **擅自新增**：合併檔不能出現來源沒有的 `(gvkey, datadate)`。
4. **去重**：每個 `(gvkey, datadate)` 只能有一列。報告會列出 NA 與 Global 重疊時保留了哪一邊。
   去重的鍵應該用 `(gvkey, datadate)`，不要用 `(gvkey, fyear)`，否則變更會計年度的公司會被誤刪。
5. **數值**：合併檔的每一列，都要能在 NA 或 Global 找到逐欄完全相同的列。
   比對時已排除日期寫法、前導零、浮點雜訊這類純格式差異。
6. **格式與來源健檢**：
   - 前導零是否遺失
   - 是否剛好 1,048,575 列（Excel 單一工作表上限，代表檔案可能被截斷）
   - SQL 的篩選條件：consol=C、indfmt=INDL、datafmt=STD/HIST_STD、costat=A/I、NA 的 curcd=USD/CAD
   - datadate 是否在 2008-01-01 ~ 2025-12-31 之間

## 輸出

合併檔所在資料夾會產生 `merge_check_report_<時間>\`，內容如下：

| 檔案 | 內容 |
|---|---|
| `報告.txt` | 結論與各項檢查結果 |
| `1_缺漏_來源有但合併檔沒有.csv` | 合併檔中缺少的來源列 |
| `2_新增_來源沒有的列.csv` | 合併檔多出、來源沒有的列 |
| `3_未去重_同鍵多列.csv` | 同一鍵出現多列的資料 |
| `4_NA與Global重疊鍵.csv` | 兩邊都有的公司年度，以及合併檔保留了哪一邊 |
| `5_數值不一致明細.csv` | 逐欄列出合併檔的值與 NA、Global 的值 |
| `6_年度對照.csv` | 各年度預期列數與實際列數 |

結束代碼：0 代表通過（可能附警告），1 代表不通過。
