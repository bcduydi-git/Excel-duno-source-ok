"""Các thay đổi tối thiểu áp lên bản sao file nguồn (dùng chung cho engine và builder)."""
import copy
import datetime as dt

# Ngày trả lãi cố định theo ngân hàng — trước đây viết cứng trong query Power BI (q_LichTraNo).
NGAY_TRA_LAI = {"LPBank": 25, "HDBank": 20, "VPBank": 15}

THAM_SO = [
    ("MocTheoDoiLichDaiHan", dt.date(2026, 9, 1),
     "Khoản chỉ có lịch ở LICH_DAI_HAN: dư nợ lấy từ dòng lịch đầu tiên từ ngày này, chỉ trừ giao dịch từ ngày này."),
    ("NguongLechLai", 5000000,
     "Kỳ đã có giao dịch lãi: thiếu so với lãi dự kiến không quá mức này vẫn coi là đã trả (quy tắc Power BI)."),
]

THU_MUC_SAO_KE = [
    ("HP", "VPBank", r"C:\DuNo\SaoKe\HP\VPBank"),
    ("HP", "NABank", r"C:\DuNo\SaoKe\HP\NABank"),
    ("DNL", "NABank", r"C:\Users\BCDDi\OneDrive\99. DƯ NỢ CÁC CÔNG TY\SaoKe\DNL\NABank"),
]


def doc_bang_tra_no(path_md):
    """Bảng giao dịch người dùng gửi (markdown). Bỏ dòng NABank (người dùng xác nhận là rác)."""
    out = []
    for line in open(path_md, encoding="utf-8"):
        p = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(p) < 8 or p[7] == "NABank":
            continue
        n = lambda s: float(s.replace(",", "")) if s else None
        goc, lai, phi, phat = n(p[2]), n(p[3]), n(p[4]), n(p[5])
        lp = None if (lai is None and phi is None and phat is None) else (lai or 0) + (phi or 0) + (phat or 0)
        out.append(dict(NgayGiaoDich=dt.datetime.strptime(p[1], "%d/%m/%Y").date(), MaKheUoc=p[0],
                        GocDaTra=goc, LaiPhiDaTra=lp, GhiChu=p[6] or None))
    return out


def apply_changes(tables, tra_no_rows):
    t = copy.deepcopy(tables)
    nh = t["tblNganHang"]
    nh["cols"].append("NgayTraLai")
    for r in nh["rows"]:
        r["NgayTraLai"] = NGAY_TRA_LAI.get(r.get("MaNganHang"))
    for r in t["tblKheUoc"]["rows"]:
        if r.get("MaKheUoc") == "Thấu chi HDBank" and not r.get("NgayTraLaiDauTien"):
            r["NgayTraLaiDauTien"] = dt.date(2025, 10, 15)   # thay cho ngoại lệ viết cứng "ngày 15" trong query cũ
    tn = t["tblTraNoNhapTay"]
    start = tn["rows"][0]["__row"]
    tn["rows"] = [dict(r, __row=start + i) for i, r in enumerate(tra_no_rows)]
    t["tblThamSo"] = {"cols": ["ThamSo", "GiaTri", "GhiChu"],
                      "rows": [dict(ThamSo=a, GiaTri=b, GhiChu=c) for a, b, c in THAM_SO]}
    t["tblThuMucSaoKe"] = {"cols": ["MaCongTy", "MaNganHang", "ThuMuc"],
                           "rows": [dict(MaCongTy=a, MaNganHang=b, ThuMuc=c) for a, b, c in THU_MUC_SAO_KE]}
    return t
