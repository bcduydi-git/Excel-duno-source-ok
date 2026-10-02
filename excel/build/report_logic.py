"""Logic báo cáo viết bằng Python, tương đương công thức trên các sheet báo cáo.
Dùng để kiểm chứng kết quả LibreOffice/Excel và đối chiếu với Power BI cũ."""
import datetime as dt
from collections import defaultdict


class Idx:
    def __init__(self, R):
        self.R = R
        self.ku = {k["MaKheUoc"]: k for k in R["ku"]}
        self.goc = defaultdict(list)    # ma -> [(ngay, goc)]  (TinhDuNo = 1)
        self.lai_ky = defaultdict(list)  # maky -> [(ngay, laiphi)]
        for g in R["gd"]:
            if g["TinhDuNo"] == 1:
                self.goc[g["MaKheUoc"]].append((g["NgayGiaoDich"], g["GocDaTra"]))
            if g["MaKy"]:
                self.lai_ky[g["MaKy"]].append((g["NgayGiaoDich"], g["LaiPhiDaTra"]))

    def goc_da_tra(self, ma, den, tu=None):
        return sum(s for d, s in self.goc[ma] if d <= den and (tu is None or d >= tu))

    def du_no(self, ma, ngay):
        k = self.ku[ma]
        if k["NgayBatDau"] is None or k["NgayBatDau"] > ngay or k["GocBanDau"] is None:
            return 0.0
        return k["GocBanDau"] - self.goc_da_tra(ma, ngay)

    def ky(self, r, den, nguong, hom_nay):
        """Trạng thái một kỳ tính đến ngày `den` (quy tắc Power BI + 2 chỉnh sửa đã thống nhất)."""
        moc = min(den, hom_nay)
        goc_tra = self.goc_da_tra(r["MaKheUoc"], den)
        gcon = max(0.0, min(r["GocPhaiTra"], r["GocLuyKe"] - goc_tra))
        lt = [s for d, s in self.lai_ky[r["MaKy"]] if d <= den]
        lai_tra = sum(lt)
        so_gd = sum(1 for s in lt if s > 0)
        tt = self.ku[r["MaKheUoc"]]["NgayTatToan"]
        tat = tt is not None and tt <= den   # khoản đã tất toán: ngân hàng đã thu đủ gốc + lãi
        dong = tat
        lai = r["LaiPhiDuKien"] or 0.0
        if r["TruocDuLieu"]:
            lai_tra = 0.0
        if dong:
            gcon, lcon = 0.0, 0.0
        elif r["TruocDuLieu"]:
            lcon = 0.0
        elif r["GocPhaiTra"] > 0 and gcon == 0:
            lcon = 0.0
        elif lai <= 0:
            lcon = 0.0
        elif so_gd == 0:
            lcon = lai
        else:
            lcon = lai - lai_tra if lai - lai_tra > nguong else 0.0
        if tat and r["NgayDenHan"] > tt:
            st = "Đã tất toán"
        elif tat:
            st = "Đã trả"
        elif gcon + lcon == 0:
            st = "Đã trả"
        elif r["NgayDenHan"] <= moc:
            st = "Chưa trả"
        else:
            st = "Chưa đến hạn"
        return dict(GocDaTra=r["GocPhaiTra"] - gcon, LaiDaTra=lai_tra, GocCon=gcon, LaiCon=lcon, TrangThai=st)
