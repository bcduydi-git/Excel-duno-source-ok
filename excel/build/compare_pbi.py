"""Mô phỏng các measure chính của Power BI cũ (DAX trong Bao_cao_du_no.SemanticModel) trên cùng dữ liệu,
để đối chiếu với file Excel mới. Không thay thế việc so với số liệu Power BI thật."""
import datetime as dt
from collections import defaultdict

D1900 = dt.date(1900, 1, 1)


def giao_dich_pbi(tables, R):
    """q_TraNo cũ: nhập tay + sao kê VPBank/NABank (GOC/LAI_PHI), KHÔNG chống trùng."""
    from source_data import as_date, as_num
    ma_hop_le = {k["MaKheUoc"] for k in R["ku"]}
    gd = []
    for r in tables["tblTraNoNhapTay"]["rows"]:
        if r.get("MaKheUoc") in ma_hop_le and r.get("NgayGiaoDich"):
            gd.append(dict(Ma=r["MaKheUoc"], Ngay=as_date(r["NgayGiaoDich"]), Goc=as_num(r.get("GocDaTra")),
                           Lai=as_num(r.get("LaiPhiDaTra"))))
    for s in R["sk"]:
        if s["PhanLoai"] in ("GOC", "LAI_PHI") and s["MaKheUoc"]:
            gd.append(dict(Ma=s["MaKheUoc"], Ngay=s["NgayGiaoDich"], Goc=s["GocDaTra"], Lai=s["LaiPhiDaTra"]))
    return gd


def du_no_pbi(R, gd, ngay):
    out = {}
    for k in R["ku"]:
        moc = k["NgayBatDauTheoDoi"]
        tra = sum(g["Goc"] for g in gd if g["Ma"] == k["MaKheUoc"] and g["Ngay"] <= ngay and (moc is None or g["Ngay"] >= moc))
        if moc:
            out[k["MaKheUoc"]] = (k["DuNoBanDau"] - tra) if ngay >= moc else 0.0
        else:
            out[k["MaKheUoc"]] = (k["SoTienVay"] - tra) if (k["NgayNhanNo"] and k["NgayNhanNo"] <= ngay) else 0.0
    return out


def lich_pbi(R):
    """q_LichTraNo cũ: cùng ngày kỳ, lãi dự kiến = số tiền vay × lãi suất × số ngày / cơ sở (không trừ gốc trả trước)."""
    ku = {k["MaKheUoc"]: k for k in R["ku"]}
    out = []
    for r in R["ltn"]:
        r = dict(r)
        k = ku[r["MaKheUoc"]]
        if r["NguonLich"] == "TU_SINH" and k["LaiSuatNam"] is not None:
            r["LaiPBI"] = float(int(k["SoTienVay"] * k["LaiSuatNam"] * r["SoNgayTinhLai"] / k["CoSoTinhLai"] + 0.5))
        else:
            r["LaiPBI"] = r["LaiPhiDuKien"]
        out.append(r)
    return out


def ky_pbi(R, gd, lich, tu, den, nguong=5_000_000):
    """GocConPhaiTraCuoiKy, LaiPhiConPhaiTraCuoiKy, TrangThaiThanhToanKy của Power BI cũ cho các kỳ trong [tu, den]."""
    ku = {k["MaKheUoc"]: k for k in R["ku"]}
    theo_ma = defaultdict(list)
    for r in lich:
        theo_ma[r["MaKheUoc"]].append(r)
    kq = []
    for r in lich:
        if not (tu <= r["NgayDenHan"] <= den):
            continue
        ma = r["MaKheUoc"]
        bd = ku[ma]["NgayBatDauTheoDoi"] or D1900
        truoc = [x for x in theo_ma[ma] if bd <= x["NgayDenHan"] < r["NgayDenHan"]]
        goc_truoc = sum(x["GocPhaiTra"] for x in truoc)
        da_tra = sum(g["Goc"] for g in gd if g["Ma"] == ma and bd <= g["Ngay"] <= den)
        gcon = max(0.0, r["GocPhaiTra"] - max(0.0, da_tra - goc_truoc)) if (r["GocPhaiTra"] > 0 and r["NgayDenHan"] >= bd) else 0.0
        lai_truoc = [x["NgayDenHan"] for x in truoc if (x["LaiPBI"] or 0) > 0]
        moc = (max(lai_truoc) + dt.timedelta(days=1)) if lai_truoc else bd
        w = [g for g in gd if g["Ma"] == ma and moc <= g["Ngay"] <= den]
        lai_w = sum(g["Lai"] for g in w)
        co_gd = any(g["Lai"] > 0 for g in w)
        lai = r["LaiPBI"] or 0.0
        lcon = 0.0
        if lai > 0 and r["NgayDenHan"] >= bd and not (r["GocPhaiTra"] > 0 and gcon == 0):
            lcon = lai if not co_gd else (max(0.0, lai - lai_w) if max(0.0, lai - lai_w) > nguong else 0.0)
        kq.append(dict(r, GocConPBI=gcon, LaiConPBI=lcon,
                       TrangThaiPBI="Chưa thanh toán" if (gcon > 0 or lcon > 0) else "Đã thanh toán"))
    return kq
