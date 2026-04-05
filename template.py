"""
Script điền dữ liệu vào template Dự Toán Đặt Hàng
Steps 4-8 theo tài liệu mapping:
  Step 4: Làm sạch dữ liệu AMIS (xóa khách HỘ KINH DOANH NGUYỄN THỊ KHIÊM NHƯ)
  Step 5: Mở template
  Step 6: Điền SL Bán 3 tháng (AMIS 3m + ESHOP xuất kho 3m)
  Step 7: Điền SL Bán 6 tháng (AMIS 6m + ESHOP xuất kho 6m)
  Step 8: Điền Tồn kho (AMIS tồn + ESHOP cuối kỳ)
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

# ─── CONFIG ───────────────────────────────────────────────────────────────────
CUSTOMER_EXCLUDE = "HỘ KINH DOANH NGUYỄN THỊ KHIÊM NHƯ"

AMIS_3M_PATH        = "input/AMIS/So_chi_tiet_ban_hang_AMIS_3m.xlsx"
AMIS_6M_PATH        = "input/AMIS/So_chi_tiet_ban_hang_AMIS_6m.xlsx"
AMIS_TON_PATH       = "input/AMIS/Tong_hop_ton_tren_nhieu_kho_AMIS_912026.xlsx"
ESHOP_TON_3M_PATH   = "input/ESHOP/TONG_HOP_TON_KHO_eShop_3m.xlsx"
ESHOP_TON_6M_PATH   = "input/ESHOP/TONG_HOP_TON_KHO_eShop_6m.xlsx"
TEMPLATE_PATH       = "input/TEMPLATE/DU_DOAN_DAT_HANG_6_THANG_NO_DATA.xls"
OUTPUT_PATH         = "output/DU_DOAN_DAT_HANG_OUTPUT.xls"

TEMPLATE_SHEET = "VAC 6 THANG 09.07.25-09.01.26"
HEADER_ROW     = 5   # 0-indexed, row 5 chứa header (Code, Name, ...)
DATA_START_ROW = 6   # 0-indexed, data bắt đầu từ row 6

# Cột trong template (0-indexed)
COL_CODE      = 1   # Code / Mã hàng
COL_NAME      = 2   # Tên hàng
COL_SL3M      = 4   # SL BÁN 3 THÁNG
COL_SL6M      = 5   # SL BÁN 6 THÁNG
COL_TON       = 6   # TỒN KHO
COL_BQ        = 10  # BQ BÁN/NGÀY
COL_NGAY_TON  = 11  # NGÀY TỒN
COL_FORECAST  = 22  # Order Forecast

DAYS_6M = 180

# ─── STEP 4: Load & làm sạch AMIS ────────────────────────────────────────────
print("Đang load dữ liệu AMIS...")

amis_3m = pd.read_excel(AMIS_3M_PATH, header=3)
amis_6m = pd.read_excel(AMIS_6M_PATH, header=3)

for name, df in [("amis_3m", amis_3m), ("amis_6m", amis_6m)]:
    before = len(df)
    mask = df["Tên khách hàng"] == CUSTOMER_EXCLUDE
    if mask.any():
        print(f"  Xóa {mask.sum()} rows khách '{CUSTOMER_EXCLUDE}' trong {name}")

amis_3m = amis_3m[amis_3m["Tên khách hàng"] != CUSTOMER_EXCLUDE].reset_index(drop=True)
amis_6m = amis_6m[amis_6m["Tên khách hàng"] != CUSTOMER_EXCLUDE].reset_index(drop=True)

# ─── Tổng SL bán theo Mã hàng từ AMIS ────────────────────────────────────────
def sum_amis_by_sku(df):
    """Tổng 'Tổng số lượng bán' theo 'Mã hàng'"""
    col_sl = "Tổng số lượng bán"
    col_ma = "Mã hàng"
    return df.groupby(col_ma)[col_sl].sum()

amis_sl_3m = sum_amis_by_sku(amis_3m)
amis_sl_6m = sum_amis_by_sku(amis_6m)
print(f"  AMIS 3m: {len(amis_sl_3m)} SKUs | AMIS 6m: {len(amis_sl_6m)} SKUs")

# ─── Lookup tên hàng từ AMIS ──────────────────────────────────────────────────
ten_hang_map = (
    amis_3m.drop_duplicates("Mã hàng")
    .set_index("Mã hàng")["Tên hàng"]
    .to_dict()
)

# ─── Load ESHOP tồn kho 3 tháng → lấy cột Xuất kho (dùng cho Step 6) ────────
print("Đang load dữ liệu ESHOP tồn kho 3 tháng...")
eshop_ton_3m = pd.read_excel(ESHOP_TON_3M_PATH, header=3)
eshop_ton_3m = eshop_ton_3m[eshop_ton_3m["Mã hàng hóa"].notna()].reset_index(drop=True)
eshop_ton_3m = eshop_ton_3m[eshop_ton_3m["Mã hàng hóa"] != "(2)"].reset_index(drop=True)
eshop_xuat_kho_3m = eshop_ton_3m.set_index("Mã hàng hóa")["Xuất kho"]
print(f"  ESHOP tồn kho 3m: {len(eshop_xuat_kho_3m)} SKUs")

# ─── Load ESHOP tồn kho 6 tháng → lấy cột Xuất kho + Cuối kỳ ────────────────
print("Đang load dữ liệu ESHOP tồn kho 6 tháng...")
eshop_ton_6m = pd.read_excel(ESHOP_TON_6M_PATH, header=3)
eshop_ton_6m = eshop_ton_6m[eshop_ton_6m["Mã hàng hóa"].notna()].reset_index(drop=True)
eshop_ton_6m = eshop_ton_6m[eshop_ton_6m["Mã hàng hóa"] != "(2)"].reset_index(drop=True)
eshop_xuat_kho_6m = eshop_ton_6m.set_index("Mã hàng hóa")["Xuất kho"]
eshop_cuoi_ky     = eshop_ton_6m.set_index("Mã hàng hóa")["Cuối kỳ"]
print(f"  ESHOP tồn kho 6m: {len(eshop_xuat_kho_6m)} SKUs")

# Bổ sung tên hàng từ ESHOP cho các mã chưa có trong AMIS
for _, row in eshop_ton_6m.iterrows():
    ma = str(row["Mã hàng hóa"]).strip()
    if ma and ma not in ten_hang_map:
        ten_hang_map[ma] = str(row["Tên hàng hóa"]).strip()

# ─── Load AMIS tồn kho ────────────────────────────────────────────────────────
print("Đang load dữ liệu AMIS tồn kho...")
amis_ton = pd.read_excel(AMIS_TON_PATH, header=3)
amis_ton_by_sku = amis_ton.set_index("Mã hàng")["Tổng"]
print(f"  AMIS tồn kho: {len(amis_ton_by_sku)} SKUs")

# ─── Helper: tính data cho 1 mã hàng ─────────────────────────────────────────
def calc_data_for_code(code):
    # Step 6: SL Bán 3 tháng
    sl_amis_3m_val = amis_sl_3m.get(code, 0)
    sl_eshop_xk_3m = eshop_xuat_kho_3m.get(code, 0)
    if sl_eshop_xk_3m == 0:
        matches = [v for k, v in eshop_xuat_kho_3m.items()
                   if str(k).startswith(code) or code.startswith(str(k).split("-")[0])]
        sl_eshop_xk_3m = sum(matches)
    sl_3m = sl_amis_3m_val + sl_eshop_xk_3m

    # Step 7: SL Bán 6 tháng
    sl_amis_6m_val = amis_sl_6m.get(code, 0)
    sl_eshop_xk_6m = eshop_xuat_kho_6m.get(code, 0)
    if sl_eshop_xk_6m == 0:
        matches = [v for k, v in eshop_xuat_kho_6m.items()
                   if str(k).startswith(code) or code.startswith(str(k).split("-")[0])]
        sl_eshop_xk_6m = sum(matches)
    sl_6m = sl_amis_6m_val + sl_eshop_xk_6m

    # Step 8: Tồn kho
    ton_amis  = amis_ton_by_sku.get(code, 0)
    ton_eshop = eshop_cuoi_ky.get(code, 0)
    if ton_eshop == 0:
        matches_ton = [v for k, v in eshop_cuoi_ky.items()
                       if str(k).startswith(code) or code.startswith(str(k).split("-")[0])]
        ton_eshop = sum(matches_ton)
    ton_kho = ton_amis + ton_eshop

    # Step 9: BQ Bán/Ngày
    bq_ban_ngay = sl_6m / DAYS_6M if sl_6m > 0 else 0

    # Step 10: Ngày Tồn
    ngay_ton = (ton_kho / bq_ban_ngay) if bq_ban_ngay > 0 else 99999

    # Step 12: Order Forecast
    if ngay_ton < DAYS_6M:
        order_forecast = round((DAYS_6M - ngay_ton) * bq_ban_ngay)
    else:
        order_forecast = 0

    return sl_3m, sl_6m, ton_kho, bq_ban_ngay, ngay_ton, order_forecast


# ─── STEP 5: Mở template, đọc mã hàng đã có sẵn ─────────────────────────────
print(f"\nĐang mở template: {TEMPLATE_PATH}")
rb = xlrd.open_workbook(TEMPLATE_PATH, formatting_info=True)

# Patch: một số format_str bị None trong file xls → xlwt crash khi save
for fmt in rb.format_map.values():
    if fmt.format_str is None:
        fmt.format_str = "General"

wb = xl_copy(rb)

sheet_names = rb.sheet_names()
sheet_idx = sheet_names.index(TEMPLATE_SHEET)
ws = wb.get_sheet(sheet_idx)
rs = rb.sheet_by_index(sheet_idx)

# Đọc tất cả mã hàng đã có trong template (giữ nguyên thứ tự, không xóa)
template_codes = {}   # code -> row_idx
for row_idx in range(DATA_START_ROW, rs.nrows):
    row_vals = rs.row_values(row_idx)
    code = str(row_vals[COL_CODE]).strip()
    if code and code not in ("", "nan"):
        template_codes[code] = row_idx

print(f"  Template có {len(template_codes)} mã hàng sẵn có.")

# ─── STEP 6-8: Điền data cho mã đã có trong template ─────────────────────────
updated = 0
for stt, (code, row_idx) in enumerate(template_codes.items(), start=1):
    sl_3m, sl_6m, ton_kho, bq_ban_ngay, ngay_ton, order_forecast = calc_data_for_code(code)

    ws.write(row_idx, 0,            stt)
    ws.write(row_idx, COL_NAME,     ten_hang_map.get(code, ""))
    ws.write(row_idx, COL_SL3M,     float(sl_3m))
    ws.write(row_idx, COL_SL6M,     float(sl_6m))
    ws.write(row_idx, COL_TON,      float(ton_kho))
    ws.write(row_idx, COL_BQ,       round(float(bq_ban_ngay), 4))
    ws.write(row_idx, COL_NGAY_TON, round(float(ngay_ton), 2))
    ws.write(row_idx, COL_FORECAST, float(order_forecast))
    updated += 1

print(f"  Đã cập nhật {updated} mã hàng có sẵn trong template.")

# ─── Thu thập tất cả mã từ AMIS + ESHOP, insert mã mới vào template ──────────
all_source_codes = set()
all_source_codes.update(str(k).strip() for k in amis_sl_3m.index)
all_source_codes.update(str(k).strip() for k in amis_sl_6m.index)
all_source_codes.update(str(k).strip() for k in amis_ton_by_sku.index)
all_source_codes.update(str(k).strip() for k in eshop_xuat_kho_3m.index)
all_source_codes.update(str(k).strip() for k in eshop_xuat_kho_6m.index)
all_source_codes.update(str(k).strip() for k in eshop_cuoi_ky.index)
all_source_codes = {c for c in all_source_codes if c and c not in ("", "nan")}

new_codes = sorted(all_source_codes - set(template_codes.keys()))
print(f"  Phát hiện {len(new_codes)} mã hàng mới (chưa có trong template), sẽ insert thêm.")

next_row = rs.nrows
inserted = 0
for code in new_codes:
    sl_3m, sl_6m, ton_kho, bq_ban_ngay, ngay_ton, order_forecast = calc_data_for_code(code)

    stt = updated + inserted + 1
    ws.write(next_row, 0,            stt)
    ws.write(next_row, COL_CODE,     code)
    ws.write(next_row, COL_NAME,     ten_hang_map.get(code, ""))
    ws.write(next_row, COL_SL3M,     float(sl_3m))
    ws.write(next_row, COL_SL6M,     float(sl_6m))
    ws.write(next_row, COL_TON,      float(ton_kho))
    ws.write(next_row, COL_BQ,       round(float(bq_ban_ngay), 4))
    ws.write(next_row, COL_NGAY_TON, round(float(ngay_ton), 2))
    ws.write(next_row, COL_FORECAST, float(order_forecast))
    next_row += 1
    inserted += 1

print(f"\nTổng kết: cập nhật {updated} mã có sẵn, insert thêm {inserted} mã mới.")

# ─── Lưu file output ──────────────────────────────────────────────────────────
import os
os.makedirs("output", exist_ok=True)
wb.save(OUTPUT_PATH)
print(f"Đã lưu file: {OUTPUT_PATH}")
