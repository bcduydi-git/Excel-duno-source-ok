"""Dựng workbook báo cáo dư nợ từ BẢN SAO file nhập liệu.

    python build_workbook.py <file_nguon.xlsx> <bang_tra_no.md> <file_ket_qua.xlsx> [--sao-ke HP:VPBank:file ...]

- Giữ nguyên các sheet nhập liệu; chỉ thêm: cột NgayTraLai (DANH_MUC), sheet CAU_HINH.
- Thêm bảng kết quả Power Query (DL_*, NGOAI_LE) với dữ liệu ban đầu tính bằng engine.py,
  4 sheet báo cáo bằng công thức, rồi nhúng các query (pq_inject.py).
"""
import argparse
import datetime as dt
import os
import sys

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

sys.path.insert(0, os.path.dirname(__file__))
import engine  # noqa: E402
import expected  # noqa: E402
import formulas as F  # noqa: E402
import pq_inject  # noqa: E402
from prepare import THAM_SO, THU_MUC_SAO_KE, apply_changes, doc_bang_tra_no  # noqa: E402
from source_data import read_tables  # noqa: E402

PQ_DIR = os.path.join(os.path.dirname(__file__), "..", "power_query")

DATE = "dd/mm/yyyy"
MONEY = "#,##0;-#,##0;-"
PCT = "0.00%"
TY = '#,##0.00,,," tỷ";-#,##0.00,,," tỷ";"-"'
INPUT = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="1F4E78")
TOTAL = PatternFill("solid", fgColor="DDEBF7")
KPI = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")

# Bảng kết quả Power Query: tên query = tên bảng, sheet, cột (theo đúng thứ tự cột của query)
KQ = [
    ("KQ_KheUoc", "DL_KheUoc", ["MaCongTy", "TenCongTy", "MaNganHang", "MaKheUoc", "MaHDTD", "NgayNhanNo", "NgayDaoHan",
                                "SoTienVay", "LaiSuatNam", "PhuongThucTraGoc", "TanSuatTraLai", "NgayTraLaiDauTien",
                                "CoSoTinhLai", "DuNoBanDau", "NgayBatDauTheoDoi", "GocBanDau", "NgayBatDau",
                                "NgayTatToan", "GhiChu"]),
    ("KQ_LichTraNo", "DL_LichTraNo", ["MaCongTy", "MaNganHang", "MaKheUoc", "MaKy", "KyTraNo", "TuNgayTinhLai",
                                      "NgayDenHan", "SoNgayTinhLai", "GocPhaiTra", "LaiPhiDuKien", "TongPhaiTra",
                                      "GocLuyKe", "TruocDuLieu", "NguonLich"]),
    ("KQ_GiaoDich", "DL_GiaoDich", ["Nguon", "IDGiaoDich", "MaCongTy", "MaNganHang", "MaHDTD", "MaKheUoc",
                                    "NgayGiaoDich", "GocDaTra", "LaiPhiDaTra", "TinhDuNo", "MaKy", "NgayDenHanKy",
                                    "NoiDung", "TenFileNguon", "DongNguon"]),
    ("KQ_NgoaiLe", "NGOAI_LE", ["Loai", "MaKheUoc", "NgayGiaoDich", "SoTien", "ChiTiet", "Nguon", "TenFileNguon",
                                "DongNguon"]),
]
KQ_KEY = {"KQ_KheUoc": "ku", "KQ_LichTraNo": "ltn", "KQ_GiaoDich": "gd", "KQ_NgoaiLe": "nl"}
# Query không tải ra sheet (chỉ kết nối)
PRIVATE = ["ThamSo", "fnSaoKeVPBank", "fnSaoKeNABank", "SaoKe", "DanhMucKheUoc", "LichKhung", "GiaoDichNguon"]


def fmt_for(col):
    if col.startswith("Ngay") or col.startswith("TuNgay"):
        return DATE
    if col in ("LaiSuatNam",):
        return PCT
    if any(col.startswith(p) for p in ("SoTien", "Goc", "LaiPhi", "TongPhaiTra", "DuNo")):
        return MONEY
    return None


def ghi_bang(wb, ten, sheet, cols, rows, ghi_chu):
    ws = wb.create_sheet(sheet)
    ws["A1"] = ghi_chu
    ws["A1"].font = Font(italic=True, color="595959")
    r0 = 3
    for j, c in enumerate(cols, 1):
        ws.cell(r0, j, c)
    for i, r in enumerate(rows, 1):
        for j, c in enumerate(cols, 1):
            v = r.get(c)
            cell = ws.cell(r0 + i, j, v)
            f = fmt_for(c)
            if f:
                cell.number_format = f
    last = r0 + max(len(rows), 1)
    ref = f"A{r0}:{get_column_letter(len(cols))}{last}"
    t = Table(displayName=ten, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    for j, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = 14 if fmt_for(c) else 18
    ws.freeze_panes = f"A{r0 + 1}"
    return ws


def tieu_de(ws, title, note):
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    ws["A2"] = note
    ws["A2"].font = Font(italic=True, color="595959")


def tham_so(ws, row, label, value, fmt=None, dv=None, hint=None):
    ws.cell(row, 1, label).font = Font(bold=True)
    c = ws.cell(row, 2, value)
    c.fill = INPUT
    c.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    if fmt:
        c.number_format = fmt
    if dv is not None:
        dv.add(c)
    if hint:
        ws.cell(row, 3, hint).font = Font(italic=True, color="7F7F7F")


def kpi(ws, row, col, label, formula, fmt=TY):
    a = ws.cell(row, col, label)
    a.font = Font(bold=True, size=9, color="404040")
    a.fill = KPI
    a.alignment = Alignment(wrap_text=True, vertical="top")
    b = ws.cell(row + 1, col, formula)
    b.font = Font(bold=True, size=12)
    b.fill = KPI
    b.number_format = fmt


def header(ws, row, cols, formats, widths):
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row, j, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        L = get_column_letter(j)
        ws.column_dimensions[L].width = widths[j - 1]
        if formats[j - 1]:
            for r in range(row + 1, row + 1500):
                ws.cell(r, j).number_format = formats[j - 1]
    ws.row_dimensions[row].height = 32
    for j in range(1, len(cols) + 1):
        ws.cell(row + 1, j).font = Font(bold=True)
        ws.cell(row + 1, j).fill = TOTAL
    ws.freeze_panes = ws.cell(row + 2, 1)


def mang(ws, cell, formula, nrows, ncols):
    """Ghi công thức mảng động; vùng ref theo kích thước kết quả tính trước (Excel tự điều chỉnh khi tính lại)."""
    c = ws[cell]
    ref = f"{cell}:{get_column_letter(c.column + ncols - 1)}{c.row + nrows - 1}"
    ws[cell] = ArrayFormula(ref, F.to_file(formula))


def build(src, tra_no_md, out, sao_ke_files, hom_nay):
    tables = apply_changes(read_tables(src), doc_bang_tra_no(tra_no_md))
    R = engine.chay(tables, sao_ke_files)

    wb = openpyxl.load_workbook(src)
    # ---------------- thay đổi tối thiểu trên sheet nhập liệu
    ws = wb["DANH_MUC"]
    t = ws.tables["tblNganHang"]
    t.ref = "D5:F10"
    t.autoFilter.ref = "D5:F10" if t.autoFilter is not None else None
    t.tableColumns.append(TableColumn(id=len(t.tableColumns) + 1, name="NgayTraLai"))
    ws["F5"] = "NgayTraLai"
    ws["F5"]._style = ws["E5"]._style
    ws["F4"] = "Ngày trả lãi cố định (tự sinh lịch)"
    ws["F4"].font = Font(italic=True, size=9, color="595959")
    for r in range(6, 11):
        nh = ws.cell(r, 4).value
        ws.cell(r, 6, next((x["NgayTraLai"] for x in tables["tblNganHang"]["rows"] if x["MaNganHang"] == nh), None))
        ws.cell(r, 6)._style = ws.cell(r, 5)._style
    ws.column_dimensions["F"].width = 14

    ws = wb["KHE_UOC"]
    for r in range(6, 96):
        if ws.cell(r, 3).value == "Thấu chi HDBank" and not ws.cell(r, 13).value:
            ws.cell(r, 13, dt.date(2025, 10, 15)).number_format = DATE

    ws = wb["TRA_NO_NHAP_TAY"]
    rows = tables["tblTraNoNhapTay"]["rows"]
    for r in range(6, 340):
        for c in range(1, 6):
            ws.cell(r, c).value = None
    for i, x in enumerate(rows):
        r = 6 + i
        ws.cell(r, 1, x["NgayGiaoDich"]).number_format = DATE
        ws.cell(r, 2, x["MaKheUoc"])
        ws.cell(r, 3, x["GocDaTra"]).number_format = MONEY
        ws.cell(r, 4, x["LaiPhiDaTra"]).number_format = MONEY
        ws.cell(r, 5, x["GhiChu"])
    ws["D5"].comment = None
    ws["A3"] = ("Mỗi dòng là một lần trả / một mã khế ước. LaiPhiDaTra = mọi khoản ngoài gốc (lãi, phí, phạt). "
                "Giao dịch đã có trong sao kê tự động (VPBank, NABank) không cần nhập; nếu nhập trùng ngày sẽ tự bỏ.")

    # Khôi phục danh sách chọn (data validation dạng mở rộng mà openpyxl không giữ được)
    for sh, rng, f in [("KHE_UOC", "A6:A95", "DANH_MUC!$A$6:$A$105"), ("KHE_UOC", "B6:B95", "DANH_MUC!$D$6:$D$10"),
                       ("HOP_DONG", "A6:A10", "DANH_MUC!$A$6:$A$105"), ("HOP_DONG", "C6:C10", "DANH_MUC!$D$6:$D$105"),
                       ("LICH_DAI_HAN", "A6:A196", "DANH_MUC!$A$6:$A$105"), ("LICH_DAI_HAN", "B6:B196", "DANH_MUC!$D$6:$D$10")]:
        dv = DataValidation(type="list", formula1=f, allow_blank=True)
        dv.add(rng)
        wb[sh].add_data_validation(dv)

    # ---------------- CAU_HINH
    ws = wb.create_sheet("CAU_HINH")
    tieu_de(ws, "CẤU HÌNH", "Sửa giá trị ở bảng; thêm thư mục sao kê bằng cách thêm dòng. Sau khi sửa: Data > Refresh All.")
    ws["A3"] = "Tham số"
    ws["A3"].font = Font(bold=True)
    for j, c in enumerate(["ThamSo", "GiaTri", "GhiChu"], 1):
        ws.cell(4, j, c)
    for i, (a, b, c) in enumerate(THAM_SO, 5):
        ws.cell(i, 1, a)
        ws.cell(i, 2, b).number_format = DATE if isinstance(b, dt.date) else MONEY
        ws.cell(i, 2).fill = INPUT
        ws.cell(i, 3, c)
    ws.add_table(Table(displayName="tblThamSo", ref=f"A4:C{4 + len(THAM_SO)}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True)))
    r0 = 6 + len(THAM_SO) + 2
    ws.cell(r0 - 1, 1, "Thư mục sao kê đọc tự động (VPBank, NABank)").font = Font(bold=True)
    for j, c in enumerate(["MaCongTy", "MaNganHang", "ThuMuc"], 1):
        ws.cell(r0, j, c)
    for i, (a, b, c) in enumerate(THU_MUC_SAO_KE, r0 + 1):
        ws.cell(i, 1, a); ws.cell(i, 2, b); ws.cell(i, 3, c).fill = INPUT
    ws.add_table(Table(displayName="tblThuMucSaoKe", ref=f"A{r0}:C{r0 + len(THU_MUC_SAO_KE)}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True)))
    huong_dan = [
        "CÁCH DÙNG",
        "• Thêm khoản vay: thêm dòng ở KHE_UOC (khoản trả gốc theo lịch ngân hàng: nhập thêm lịch ở LICH_DAI_HAN). Refresh All.",
        "• Trả nợ: VPBank/NABank thả file sao kê vào thư mục ở bảng trên; ngân hàng khác (LPBank, HDBank, Chailease…) nhập ở TRA_NO_NHAP_TAY. Refresh All.",
        "• Xem báo cáo: 4 sheet 1_… 4_… — đổi ô màu vàng (ngày, pháp nhân, ngân hàng, trạng thái), kết quả cập nhật ngay, không cần Refresh.",
        "• Sau mỗi lần Refresh xem sheet NGOAI_LE: chỉ liệt kê những gì không tự xử lý được (sai mã, trùng, sao kê chưa tách gốc/lãi…).",
        "• Chi tiết đã tính nằm ở DL_KheUoc, DL_LichTraNo, DL_GiaoDich (lọc bằng nút lọc của bảng). Mỗi giao dịch có tên file + số dòng nguồn.",
        "• Lần đầu mở trên máy: Data > Get Data > Query Options > Current Workbook > Privacy: chọn 'Ignore the Privacy Levels' (để query đọc được thư mục sao kê theo cấu hình).",
    ]
    for i, s in enumerate(huong_dan, r0 + len(THU_MUC_SAO_KE) + 3):
        ws.cell(i, 1, s).font = Font(bold=(s == "CÁCH DÙNG"))
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 90

    # ---------------- bảng kết quả Power Query
    notes = {
        "KQ_KheUoc": "Kết quả Power Query – danh mục khế ước (KHE_UOC + khoản chỉ có lịch dài hạn). Không sửa tay; Refresh để cập nhật.",
        "KQ_LichTraNo": "Kết quả Power Query – lịch trả nợ (lịch ngân hàng + lịch tự sinh). Lãi/phí dự kiến chỉ là ước tính.",
        "KQ_GiaoDich": "Kết quả Power Query – giao dịch trả nợ được tính (nhập tay + sao kê, đã chống trùng). MaKy = kỳ được ghép cho giao dịch lãi/phí.",
        "KQ_NgoaiLe": "Những trường hợp cần xem sau mỗi lần Refresh. Dòng có '(không tính)' / 'đã bỏ' không được cộng vào báo cáo.",
    }
    for ten, sheet, cols in KQ:
        ghi_bang(wb, ten, sheet, cols, R[KQ_KEY[ten]], notes[ten])
    wb["NGOAI_LE"].column_dimensions["A"].width = 52
    wb["NGOAI_LE"].column_dimensions["E"].width = 70

    # ---------------- báo cáo
    dv_pn = lambda: DataValidation(type="list", formula1="DANH_MUC!$A$6:$A$105", allow_blank=True)
    dv_nh = lambda: DataValidation(type="list", formula1="DANH_MUC!$D$6:$D$105", allow_blank=True)
    tu0 = dt.date(hom_nay.year, hom_nay.month, 1)
    den0 = (tu0.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)

    # 1. Tổng quan
    ws = wb.create_sheet("1_TONG_QUAN", 0)
    tieu_de(ws, "1. TỔNG QUAN DƯ NỢ VÀ NGHĨA VỤ THANH TOÁN",
            "Dư nợ gốc theo giao dịch thực trả; nghĩa vụ = các kỳ đến hạn trong khoảng ngày, tình trạng tính đến 'Đến ngày'.")
    a, b = dv_pn(), dv_nh(); ws.add_data_validation(a); ws.add_data_validation(b)
    tham_so(ws, 3, "Từ ngày", "=EOMONTH(TODAY(),-1)+1", DATE, hint=None)
    tham_so(ws, 4, "Đến ngày", "=EOMONTH(TODAY(),0)", DATE)
    tham_so(ws, 5, "Pháp nhân", None, dv=a, hint="trống = tất cả")
    tham_so(ws, 6, "Ngân hàng", None, dv=b, hint="trống = tất cả")
    for col, lab, f in [(5, "Dư nợ đầu kỳ", "=D12"), (6, "Giải ngân trong kỳ", "=E12"), (7, "Trả gốc trong kỳ", "=F12"),
                        (8, "Dư nợ cuối kỳ", "=G12")]:
        kpi(ws, 3, col, lab, f)
    for col, lab in [(10, "Hạn mức đang hiệu lực"), (11, "Dư nợ trong hạn mức"), (12, "Dư nợ quá hạn")]:
        kpi(ws, 3, col, lab, None)
    mang(ws, "J4", F.tong_quan_han_muc(), 1, 3)
    for col, lab, f in [(5, "Gốc đến hạn trong kỳ", "=H12"), (6, "Lãi/phí dự kiến đến hạn", "=I12"),
                        (7, "Lãi/phí đã trả trong kỳ", "=J12"), (8, "Gốc còn phải trả", "=K12"),
                        (9, "Lãi/phí còn phải trả", "=L12"), (10, "Hạn mức chưa sử dụng", "=J4-K4")]:
        kpi(ws, 6, col, lab, f)
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[6].height = 30
    header(ws, 11, F.TONG_QUAN_COLS, [None, None, "#,##0"] + [MONEY] * 9, [12, 12, 10] + [17] * 9)
    exp = expected.tong_quan(R, tu0, den0, hom_nay=hom_nay)
    mang(ws, "A12", F.tong_quan(), len(exp), len(F.TONG_QUAN_COLS))

    # 2. HĐTD & hạn mức
    ws = wb.create_sheet("2_HDTD_HAN_MUC", 1)
    tieu_de(ws, "2. HỢP ĐỒNG TÍN DỤNG VÀ HẠN MỨC",
            "Dư nợ thuộc HĐ = dư nợ các khế ước cùng pháp nhân + mã HĐTD. Hạn mức chưa sử dụng chỉ tính cho HĐ đang hiệu lực.")
    a, b = dv_pn(), dv_nh(); ws.add_data_validation(a); ws.add_data_validation(b)
    tham_so(ws, 3, "Ngày báo cáo", "=TODAY()", DATE)
    tham_so(ws, 4, "Pháp nhân", None, dv=a, hint="trống = tất cả")
    tham_so(ws, 5, "Ngân hàng", None, dv=b, hint="trống = tất cả")
    for col, lab, f, fm in [(5, "Hạn mức đang hiệu lực", "=G9", TY), (6, "Dư nợ trong hạn mức", "=H9", TY),
                            (7, "Hạn mức chưa sử dụng", "=I9", TY), (8, "Tỷ lệ sử dụng", "=J9", PCT)]:
        kpi(ws, 3, col, lab, f, fm)
    kpi(ws, 3, 9, "Dư nợ ngoài hạn mức hiệu lực (HĐ hết hạn, Chailease…)", None)
    ws.row_dimensions[3].height = 42
    mang(ws, "I4", F.ngoai_han_muc(), 1, 1)
    header(ws, 8, F.HOP_DONG_COLS, [None, None, None, DATE, DATE, None, MONEY, MONEY, MONEY, PCT, "#,##0", None, None],
           [10, 14, 11, 12, 12, 13, 18, 18, 18, 10, 9, 16, 40])
    exp = expected.hop_dong(R, tables, hom_nay)
    mang(ws, "A9", F.hop_dong(), len(exp), len(F.HOP_DONG_COLS))

    # 3. Khế ước
    ws = wb.create_sheet("3_KHE_UOC", 2)
    tieu_de(ws, "3. DANH SÁCH KHẾ ƯỚC",
            "Dư nợ, gốc và lãi/phí đã trả tính đến 'Ngày báo cáo'. Mốc tính hạn = ngày báo cáo nhưng không quá hôm nay.")
    a, b = dv_pn(), dv_nh(); ws.add_data_validation(a); ws.add_data_validation(b)
    c = DataValidation(type="list", formula1='"Còn dư nợ,Quá hạn,Đến hạn ≤ 7 ngày,Đến hạn ≤ 30 ngày,Trong hạn,Đã tất toán,Chưa giải ngân,Thiếu ngày đáo hạn"', allow_blank=True)
    ws.add_data_validation(c)
    tham_so(ws, 3, "Ngày báo cáo", "=TODAY()", DATE)
    tham_so(ws, 4, "Pháp nhân", None, dv=a, hint="trống = tất cả")
    tham_so(ws, 5, "Ngân hàng", None, dv=b, hint="trống = tất cả")
    tham_so(ws, 6, "Trạng thái", "Còn dư nợ", dv=c, hint="trống = tất cả khế ước")
    header(ws, 8, F.KHE_UOC_COLS, [None, None, None, None, DATE, DATE, MONEY, PCT, MONEY, MONEY, MONEY, DATE, "0", "0", None],
           [10, 11, 26, 13, 12, 12, 17, 9, 17, 17, 15, 12, 10, 10, 17])
    exp = expected.khe_uoc(R, hom_nay, hom_nay=hom_nay)
    mang(ws, "A9", F.khe_uoc(), len(exp), len(F.KHE_UOC_COLS))
    to_mau(ws, "O10:O1500", {"Quá hạn": "F8CBAD", "Đến hạn ≤ 7 ngày": "F8CBAD", "Đến hạn ≤ 30 ngày": "FFE699",
                             "Thiếu ngày đáo hạn": "FFE699"})

    # 4. Lịch trả nợ
    ws = wb.create_sheet("4_LICH_TRA_NO", 3)
    tieu_de(ws, "4. LỊCH TRẢ NỢ",
            "Kỳ đến hạn trong khoảng ngày. Đã trả/còn phải trả tính đến 'Đến ngày' theo giao dịch thực trả (không dùng lãi dự kiến thay số thực trả).")
    a, b = dv_pn(), dv_nh(); ws.add_data_validation(a); ws.add_data_validation(b)
    c = DataValidation(type="list", formula1='"Chưa trả,Chưa đến hạn,Đã trả,Đã tất toán"', allow_blank=True)
    ws.add_data_validation(c)
    tham_so(ws, 3, "Từ ngày", "=EOMONTH(TODAY(),-1)+1", DATE)
    tham_so(ws, 4, "Đến ngày", "=EOMONTH(TODAY(),0)", DATE)
    tham_so(ws, 5, "Pháp nhân", None, dv=a, hint="trống = tất cả")
    tham_so(ws, 6, "Ngân hàng", None, dv=b, hint="trống = tất cả")
    tham_so(ws, 7, "Trạng thái", None, dv=c, hint="trống = tất cả")
    header(ws, 9, F.LICH_COLS, [DATE, None, None, None, "0", MONEY, MONEY, MONEY, MONEY, DATE, MONEY, MONEY, None, None, None],
           [12, 10, 11, 26, 6, 16, 15, 16, 15, 12, 16, 15, 14, 30, 15])
    exp = expected.lich(R, tu0, den0, hom_nay=hom_nay)
    mang(ws, "A10", F.lich_tra_no(10), len(exp), len(F.LICH_COLS))
    to_mau(ws, "M11:M1500", {"Chưa trả": "F8CBAD", "Chưa đến hạn": "FFE699", "Đã trả": "C6EFCE", "Đã tất toán": "D9D9D9"})

    # thứ tự sheet: báo cáo, nhập liệu, cấu hình, kết quả
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    tmp = out + ".tmp.xlsx"
    wb.save(tmp)
    queries = {}
    for name in PRIVATE + [k[0] for k in KQ]:
        queries[name] = open(os.path.join(PQ_DIR, name + ".m"), encoding="utf-8").read()
    pq_inject.inject(tmp, out, queries, loaded={k[0]: k[2] for k in KQ})
    os.remove(tmp)
    return tables, R


def to_mau(ws, rng, mau):
    first = rng.split(":")[0]
    for txt, color in mau.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{txt}"'],
                                                       fill=PatternFill("solid", fgColor=color, bgColor=color)))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("tra_no_md"); ap.add_argument("out")
    ap.add_argument("--sao-ke", nargs="*", default=[], help="MaCongTy:MaNganHang:duong_dan_file")
    ap.add_argument("--hom-nay", default=None)
    a = ap.parse_args()
    files = {}
    for s in a.sao_ke:
        cty, nh, p = s.split(":", 2)
        files.setdefault((cty, nh), []).append(p)
    hn = dt.date.fromisoformat(a.hom_nay) if a.hom_nay else dt.date.today()
    build(a.src, a.tra_no_md, a.out, [(c, n, f) for (c, n), f in files.items()], hn)
    print("Đã tạo", a.out)
