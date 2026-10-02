"""Kết quả mong đợi của 4 báo cáo, tính bằng Python theo đúng logic công thức (để kiểm chứng LibreOffice/Excel)."""
import datetime as dt
from collections import defaultdict

from report_logic import Idx


def _hop(v, f):
    return v == "" or v is None or f == v


def lich(R, tu, den, pn="", nh="", ttf="", hom_nay=None):
    I = Idx(R)
    ng = R["ts"]["NguongLechLai"]
    rows = []
    for r in R["ltn"]:
        if not (tu <= r["NgayDenHan"] <= den and _hop(pn, r["MaCongTy"]) and _hop(nh, r["MaNganHang"])):
            continue
        k = I.ky(r, den, ng, hom_nay)
        tt = I.ku[r["MaKheUoc"]]["NgayTatToan"]
        tat = tt is not None and tt <= den
        gchu = "Trước lần trả lãi đầu tiên có dữ liệu" if r["TruocDuLieu"] else (f"Tất toán {tt:%d/%m/%Y}" if tat else "")
        ngtra = max([d for d, s in I.lai_ky[r["MaKy"]] if d <= den], default=None)
        if ttf and k["TrangThai"] != ttf:
            continue
        rows.append([r["NgayDenHan"], r["MaCongTy"], r["MaNganHang"], r["MaKheUoc"], r["KyTraNo"], r["GocPhaiTra"],
                     r["LaiPhiDuKien"] or 0.0, k["GocDaTra"], k["LaiDaTra"], ngtra, k["GocCon"], k["LaiCon"],
                     k["TrangThai"], gchu, r["NguonLich"]])
    if not rows:
        return [["Không có kỳ nào thỏa điều kiện lọc"]]
    s = lambda c: sum(x[c] for x in rows)
    tong = [f"TỔNG: {len(rows)} kỳ", "", "", "", "", s(5), s(6), s(7), s(8), "", s(10), s(11), "", "", ""]
    return [tong] + rows


def khe_uoc(R, d, pn="", nh="", ttf="Còn dư nợ", hom_nay=None):
    I = Idx(R)
    moc = min(d, hom_nay)
    rows = []
    for k in R["ku"]:
        if not (_hop(pn, k["MaCongTy"]) and _hop(nh, k["MaNganHang"])):
            continue
        ma = k["MaKheUoc"]
        gtra = I.goc_da_tra(ma, d)
        duno = I.du_no(ma, d)
        ltra = sum(g["LaiPhiDaTra"] for g in R["gd"] if g["MaKheUoc"] == ma and g["NgayGiaoDich"] <= d)
        con = duno > 0.5
        dao, nhan = k["NgayDaoHan"], k["NgayNhanNo"]
        ke = min([r["NgayDenHan"] for r in R["ltn"] if r["MaKheUoc"] == ma and r["NgayDenHan"] > moc], default=None)
        songay = (dao - moc).days if (con and dao and dao >= moc) else ""
        quahan = (moc - dao).days if (con and dao and dao < moc) else ""
        if not dao:
            st = "Thiếu ngày đáo hạn"
        elif not con:
            st = "Chưa giải ngân" if (nhan and nhan > moc) else "Đã tất toán"
        elif dao < moc:
            st = "Quá hạn"
        elif (dao - moc).days <= 7:
            st = "Đến hạn ≤ 7 ngày"
        elif (dao - moc).days <= 30:
            st = "Đến hạn ≤ 30 ngày"
        else:
            st = "Trong hạn"
        if ttf == "Còn dư nợ" and not con:
            continue
        if ttf not in ("", "Còn dư nợ") and st != ttf:
            continue
        rows.append([k["MaCongTy"], k["MaNganHang"], ma, k["MaHDTD"] or "", nhan or "", dao or "", k["GocBanDau"] or 0.0,
                     k["LaiSuatNam"] or "", gtra, duno, ltra, ke if (con and ke) else "", songay, quahan, st])
    if not rows:
        return [["Không có khế ước nào thỏa điều kiện lọc"]]
    rows.sort(key=lambda x: (x[0], x[1], x[5] or dt.date(1900, 1, 1)))
    s = lambda c: sum(x[c] for x in rows)
    tong = [f"TỔNG: {len(rows)} khế ước", "", "", "", "", "", s(6), "", s(8), s(9), s(10), "", "", "", ""]
    return [tong] + rows


def hop_dong(R, tables, d, pn="", nh=""):
    I = Idx(R)
    dn = {k["MaKheUoc"]: I.du_no(k["MaKheUoc"], d) for k in R["ku"]}
    rows = []
    for h in tables["tblHopDong"]["rows"]:
        if not h.get("MaHDTD") or not (_hop(pn, h["MaCongTy"]) and _hop(nh, h["MaNganHang"])):
            continue
        hl = h["NgayHieuLuc"].date() if h.get("NgayHieuLuc") else None
        het = h["NgayHetHan"].date() if h.get("NgayHetHan") else None
        hm = h["HanMucTong"] or 0
        kus = [k for k in R["ku"] if k["MaCongTy"] == h["MaCongTy"] and k["MaHDTD"] == h["MaHDTD"]]
        duno = sum(dn[k["MaKheUoc"]] for k in kus)
        soku = sum(1 for k in kus if dn[k["MaKheUoc"]] > 0.5)
        st = "Chưa hiệu lực" if (hl and hl > d) else ("Hết hạn" if (het and het < d) else "Đang hiệu lực")
        hieu = st == "Đang hiệu lực"
        rows.append([h["MaCongTy"], h["MaHDTD"], h["MaNganHang"], hl or "", het or "", st, hm, duno,
                     hm - duno if hieu else "", duno / hm if (hieu and hm) else "", soku, h.get("MucDichVay") or "",
                     h.get("TSDB") or ""])
    if not rows:
        return [["Không có hợp đồng nào thỏa điều kiện lọc"]]
    hmhl = sum(x[6] for x in rows if x[5] == "Đang hiệu lực")
    dnhl = sum(x[7] for x in rows if x[5] == "Đang hiệu lực")
    tong = [f"TỔNG ({len(rows)} HĐ) – hạn mức đang hiệu lực", "", "", "", "", "", hmhl, dnhl, hmhl - dnhl,
            dnhl / hmhl if hmhl else "", sum(x[10] for x in rows), "", ""]
    return [tong] + rows


def _trong_han_muc(R, tables, d, pn, nh, loc_hd):
    hl = set()
    for h in tables["tblHopDong"]["rows"]:
        if not h.get("MaHDTD"):
            continue
        if loc_hd and not (_hop(pn, h["MaCongTy"]) and _hop(nh, h["MaNganHang"])):
            continue
        a = h["NgayHieuLuc"].date() if h.get("NgayHieuLuc") else None
        b = h["NgayHetHan"].date() if h.get("NgayHetHan") else None
        if a and a <= d and (b is None or b >= d):
            hl.add((h["MaCongTy"], h["MaHDTD"]))
    return hl


def ngoai_han_muc(R, tables, d, pn="", nh=""):
    I = Idx(R)
    hl = _trong_han_muc(R, tables, d, pn, nh, False)
    return sum(I.du_no(k["MaKheUoc"], d) for k in R["ku"]
               if _hop(pn, k["MaCongTy"]) and _hop(nh, k["MaNganHang"]) and (k["MaCongTy"], k["MaHDTD"]) not in hl)


def tong_quan(R, tu, den, pn="", nh="", hom_nay=None):
    I = Idx(R)
    ng = R["ts"]["NguongLechLai"]
    nhom = defaultdict(lambda: [0] * 10)
    kus = [k for k in R["ku"] if _hop(pn, k["MaCongTy"]) and _hop(nh, k["MaNganHang"])]
    if not kus:
        return [["Không có khế ước nào thỏa điều kiện lọc"]]
    for k in kus:
        ma = k["MaKheUoc"]
        g = nhom[(k["MaCongTy"], k["MaNganHang"])]
        cuoi = I.du_no(ma, den)
        g[0] += 1 if cuoi > 0.5 else 0
        g[1] += I.du_no(ma, tu - dt.timedelta(days=1))
        g[2] += (k["SoTienVay"] or 0) if (k["NgayNhanNo"] and tu <= k["NgayNhanNo"] <= den) else 0
        g[3] += sum(s for d, s in I.goc[ma] if tu <= d <= den)
        g[4] += cuoi
    for r in R["ltn"]:
        if not (tu <= r["NgayDenHan"] <= den and _hop(pn, r["MaCongTy"]) and _hop(nh, r["MaNganHang"])):
            continue
        key = (r["MaCongTy"], r["MaNganHang"])
        if key not in nhom:
            continue
        k = I.ky(r, den, ng, hom_nay)
        g = nhom[key]
        g[5] += r["GocPhaiTra"]; g[6] += r["LaiPhiDuKien"] or 0.0; g[8] += k["GocCon"]; g[9] += k["LaiCon"]
    for g in R["gd"]:
        key = (g["MaCongTy"], g["MaNganHang"])
        if key in nhom and tu <= g["NgayGiaoDich"] <= den:
            nhom[key][7] += g["LaiPhiDaTra"]
    rows = [[a, b] + v for (a, b), v in sorted(nhom.items())]
    tong = ["TỔNG", ""] + [sum(x[c] for x in rows) for c in range(2, 12)]
    return [tong] + rows


def tong_quan_han_muc(R, tables, den, pn="", nh="", hom_nay=None):
    I = Idx(R)
    moc = min(den, hom_nay)
    hl = _trong_han_muc(R, tables, den, pn, nh, True)
    hm = 0
    for h in tables["tblHopDong"]["rows"]:
        if h.get("MaHDTD") and (h["MaCongTy"], h["MaHDTD"]) in hl:
            hm += h["HanMucTong"] or 0
    kus = [k for k in R["ku"] if _hop(pn, k["MaCongTy"]) and _hop(nh, k["MaNganHang"])]
    dn = {k["MaKheUoc"]: I.du_no(k["MaKheUoc"], den) for k in kus}
    trong = sum(dn[k["MaKheUoc"]] for k in kus if (k["MaCongTy"], k["MaHDTD"]) in hl)
    qh = sum(dn[k["MaKheUoc"]] for k in kus if k["NgayDaoHan"] and k["NgayDaoHan"] < moc and dn[k["MaKheUoc"]] > 0.5)
    return [[hm, trong, qh]]
