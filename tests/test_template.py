"""
Test suite cho template.py - kiểm tra từng bước có chạy đúng không.
Chạy: .venv\Scripts\python.exe tests/test_template.py
"""

import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(str(__file__)))))

import pandas as pd
import xlrd

# ─── PATHS (copy từ template.py) ──────────────────────────────────────────────
CUSTOMER_EXCLUDE  = "HỘ KINH DOANH NGUYỄN THỊ KHIÊM NHƯ"
AMIS_3M_PATH      = "input/AMIS/So_chi_tiet_ban_hang_AMIS_3m.xlsx"
AMIS_6M_PATH      = "input/AMIS/So_chi_tiet_ban_hang_AMIS_6m.xlsx"
AMIS_TON_PATH     = "input/AMIS/Tong_hop_ton_tren_nhieu_kho_AMIS_912026.xlsx"
ESHOP_TON_3M_PATH = "input/ESHOP/TONG_HOP_TON_KHO_eShop_3m.xlsx"
ESHOP_TON_6M_PATH = "input/ESHOP/TONG_HOP_TON_KHO_eShop_6m.xlsx"
TEMPLATE_PATH     = "input/TEMPLATE/DU_DOAN_DAT_HANG_6_THANG.xls"
TEMPLATE_SHEET    = "VAC 6 THANG 09.07.25-09.01.26"
OUTPUT_PATH       = "output/DU_DOAN_DAT_HANG_OUTPUT.xls"
DAYS_6M           = 180

PASS = "✅ PASS"
FAIL = "❌ FAIL"

def check(name, condition, detail=""):
    status = PASS if condition else FAIL
    print(f"  {status}  {name}" + (f" → {detail}" if detail else ""))
    return condition

results = []

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 1: File input tồn tại ──────────────────────────────────────────")
for label, path in [
    ("AMIS 3m",       AMIS_3M_PATH),
    ("AMIS 6m",       AMIS_6M_PATH),
    ("AMIS tồn kho",  AMIS_TON_PATH),
    ("ESHOP tồn 3m",  ESHOP_TON_3M_PATH),
    ("ESHOP tồn 6m",  ESHOP_TON_6M_PATH),
    ("Template",      TEMPLATE_PATH),
]:
    results.append(check(label, os.path.exists(path), path))

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 2: Load AMIS & lọc khách ───────────────────────────────────────")
try:
    amis_3m = pd.read_excel(AMIS_3M_PATH, header=3)
    amis_6m = pd.read_excel(AMIS_6M_PATH, header=3)

    results.append(check("AMIS 3m load OK", len(amis_3m) > 0, f"{len(amis_3m)} rows"))
    results.append(check("AMIS 6m load OK", len(amis_6m) > 0, f"{len(amis_6m)} rows"))

    results.append(check("Cột 'Tên khách hàng' tồn tại (3m)", "Tên khách hàng" in amis_3m.columns))
    results.append(check("Cột 'Tên khách hàng' tồn tại (6m)", "Tên khách hàng" in amis_6m.columns))
    results.append(check("Cột 'Mã hàng' tồn tại (3m)",        "Mã hàng" in amis_3m.columns))
    results.append(check("Cột 'Tổng số lượng bán' tồn tại (3m)", "Tổng số lượng bán" in amis_3m.columns))

    before_3m = len(amis_3m)
    amis_3m = amis_3m[amis_3m["Tên khách hàng"] != CUSTOMER_EXCLUDE].reset_index(drop=True)
    amis_6m = amis_6m[amis_6m["Tên khách hàng"] != CUSTOMER_EXCLUDE].reset_index(drop=True)
    results.append(check("Lọc khách AMIS 3m OK", len(amis_3m) <= before_3m,
                          f"{before_3m} → {len(amis_3m)} rows"))

    amis_sl_3m = amis_3m.groupby("Mã hàng")["Tổng số lượng bán"].sum()
    amis_sl_6m = amis_6m.groupby("Mã hàng")["Tổng số lượng bán"].sum()
    results.append(check("Group by SKU 3m OK", len(amis_sl_3m) > 0, f"{len(amis_sl_3m)} SKUs"))
    results.append(check("Group by SKU 6m OK", len(amis_sl_6m) > 0, f"{len(amis_sl_6m)} SKUs"))
except Exception as e:
    results.append(check("AMIS load/filter", False, str(e)))
    amis_sl_3m = amis_sl_6m = pd.Series(dtype=float)

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 3: Load ESHOP tồn kho ──────────────────────────────────────────")
eshop_xuat_kho_3m = eshop_xuat_kho_6m = eshop_cuoi_ky = pd.Series(dtype=float)
for label, path, col in [
    ("ESHOP tồn 3m - Xuất kho", ESHOP_TON_3M_PATH, "Xuất kho"),
    ("ESHOP tồn 6m - Xuất kho", ESHOP_TON_6M_PATH, "Xuất kho"),
    ("ESHOP tồn 6m - Cuối kỳ",  ESHOP_TON_6M_PATH, "Cuối kỳ"),
]:
    try:
        df = pd.read_excel(path, header=3)
        df = df[df["Mã hàng hóa"].notna()]
        df = df[df["Mã hàng hóa"] != "(2)"]
        results.append(check(f"Cột '{col}' tồn tại", col in df.columns, f"{len(df)} rows"))
        s = df.set_index("Mã hàng hóa")[col]
        if "3m" in label:
            eshop_xuat_kho_3m = s
        elif "Xuất" in label:
            eshop_xuat_kho_6m = s
        else:
            eshop_cuoi_ky = s
    except Exception as e:
        results.append(check(label, False, str(e)))

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 4: Load AMIS tồn kho ───────────────────────────────────────────")
amis_ton_by_sku = pd.Series(dtype=float)
try:
    amis_ton = pd.read_excel(AMIS_TON_PATH, header=3)
    results.append(check("AMIS tồn kho load OK", len(amis_ton) > 0, f"{len(amis_ton)} rows"))
    results.append(check("Cột 'Mã hàng' tồn tại",  "Mã hàng" in amis_ton.columns))
    results.append(check("Cột 'Tổng' tồn tại",      "Tổng" in amis_ton.columns))
    amis_ton_by_sku = amis_ton.set_index("Mã hàng")["Tổng"]
except Exception as e:
    results.append(check("AMIS tồn kho", False, str(e)))

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 5: Template mở được & sheet đúng ───────────────────────────────")
try:
    rb = xlrd.open_workbook(TEMPLATE_PATH, formatting_info=True)
    results.append(check("Template mở OK", True))
    results.append(check(f"Sheet '{TEMPLATE_SHEET}' tồn tại",
                          TEMPLATE_SHEET in rb.sheet_names(),
                          str(rb.sheet_names())))
    rs = rb.sheet_by_name(TEMPLATE_SHEET)
    results.append(check("Template có data rows", rs.nrows > 6, f"{rs.nrows} rows"))
except Exception as e:
    results.append(check("Template", False, str(e)))

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 6: Tính toán mẫu cho 1 SKU ────────────────────────────────────")
try:
    sample_code = list(amis_sl_6m.index)[0] if len(amis_sl_6m) > 0 else None
    if sample_code:
        sl_3m = amis_sl_3m.get(sample_code, 0) + eshop_xuat_kho_3m.get(sample_code, 0)
        sl_6m = amis_sl_6m.get(sample_code, 0) + eshop_xuat_kho_6m.get(sample_code, 0)
        ton   = amis_ton_by_sku.get(sample_code, 0) + eshop_cuoi_ky.get(sample_code, 0)
        bq    = sl_6m / DAYS_6M if sl_6m > 0 else 0
        ngay  = (ton / bq) if bq > 0 else 99999
        fc    = round((DAYS_6M - ngay) * bq) if ngay < DAYS_6M else 0

        results.append(check("SL 3m >= 0",  sl_3m >= 0,  f"SKU={sample_code}, SL3m={sl_3m}"))
        results.append(check("SL 6m >= 0",  sl_6m >= 0,  f"SL6m={sl_6m}"))
        results.append(check("Tồn kho >= 0", ton >= 0,   f"Tồn={ton}"))
        results.append(check("BQ >= 0",      bq >= 0,    f"BQ={round(bq,4)}"))
        results.append(check("Ngày tồn >= 0", ngay >= 0, f"NgàyTồn={round(ngay,2)}"))
        results.append(check("Forecast >= 0", fc >= 0,   f"Forecast={fc}"))
    else:
        results.append(check("Có SKU để test", False, "amis_sl_6m rỗng"))
except Exception as e:
    results.append(check("Tính toán mẫu", False, str(e)))

# ══════════════════════════════════════════════════════════════════════════════
print("\n── TEST 7: Output file tồn tại sau khi chạy template.py ───────────────")
results.append(check("Output file tồn tại", os.path.exists(OUTPUT_PATH), OUTPUT_PATH))
if os.path.exists(OUTPUT_PATH):
    try:
        rb_out = xlrd.open_workbook(OUTPUT_PATH)
        rs_out = rb_out.sheet_by_name(TEMPLATE_SHEET)
        # Kiểm tra row đầu tiên có data
        row6 = rs_out.row_values(6)
        results.append(check("Output có giá trị SL3m", row6[4] not in (None, ""), f"col4={row6[4]}"))
        results.append(check("Output có giá trị SL6m", row6[5] not in (None, ""), f"col5={row6[5]}"))
        results.append(check("Output có giá trị Tồn",  row6[6] not in (None, ""), f"col6={row6[6]}"))
    except Exception as e:
        results.append(check("Output đọc được", False, str(e)))

# ══════════════════════════════════════════════════════════════════════════════
passed = sum(results)
total  = len(results)
print(f"\n{'═'*60}")
print(f"KẾT QUẢ: {passed}/{total} tests passed")
if passed == total:
    print("🎉 TẤT CẢ TESTS ĐỀU PASS!")
else:
    print(f"⚠️  {total - passed} test(s) FAIL - kiểm tra lại các mục ❌ ở trên")
print('═'*60)
