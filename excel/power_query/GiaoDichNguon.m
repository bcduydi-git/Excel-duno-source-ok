// Mọi giao dịch trả nợ ứng viên từ 2 nguồn: nhập tay (TRA_NO_NHAP_TAY) và sao kê (SaoKe).
// LyDoLoai = null: được tính. Khác null: không tính, đưa sang KQ_NgoaiLe kèm lý do.
// Chống tính trùng: dòng nhập tay có cùng mã khế ước + ngày với giao dịch sao kê đã tính thì bỏ dòng nhập tay.
let
    Chu = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    So = (x as any) as number => if Chu(x) = "" then 0 else if x is number then x else Number.FromText(Chu(x), "vi-VN"),
    MaHopLe = List.Buffer(DanhMucKheUoc[MaKheUoc]),
    Kieu = type table [
        Nguon = text, IDGiaoDich = text, MaKheUoc = nullable text, NgayGiaoDich = nullable date,
        GocDaTra = number, LaiPhiDaTra = number, NoiDung = nullable text, TenFileNguon = text,
        DongNguon = Int64.Type, SoTienGhi = nullable number, LyDoLoai = nullable text, ChiTiet = nullable text],
    // ---- Nhập tay (dòng đầu tiên của bảng là dòng 6 trên sheet, như Power BI cũ)
    Nhap = Excel.CurrentWorkbook(){[Name = "tblTraNoNhapTay"]}[Content],
    ChonCot = Table.SelectColumns(Nhap, {"NgayGiaoDich", "MaKheUoc", "GocDaTra", "LaiPhiDaTra", "GhiChu"}),
    SoDong = Table.AddIndexColumn(ChonCot, "DongNguon", 6, 1, Int64.Type),
    BoTrong = Table.SelectRows(SoDong, each List.AnyTrue(List.Transform(
        {[NgayGiaoDich], [MaKheUoc], [GocDaTra], [LaiPhiDaTra], [GhiChu]}, (x) => Chu(x) <> ""))),
    Tay = List.Transform(Table.ToRecords(BoTrong), (r) =>
        let
            Ma = if Chu(r[MaKheUoc]) = "" then null else Chu(r[MaKheUoc]),
            Ngay = if Chu(r[NgayGiaoDich]) = "" then null else try Date.From(r[NgayGiaoDich], "vi-VN") otherwise null,
            ThieuTien = Chu(r[GocDaTra]) = "" and Chu(r[LaiPhiDaTra]) = "",
            GhiChu = if Chu(r[GhiChu]) = "" then null else Chu(r[GhiChu]),
            Thieu = Ma = null or Ngay = null or ThieuTien,
            LyDo = if Thieu then "Thiếu thông tin ở dòng nhập tay (không tính)"
                else if not List.Contains(MaHopLe, Ma) then "Mã khế ước không có trong danh mục (không tính)"
                else null
        in [
            Nguon = "NHAP_TAY", IDGiaoDich = "NHAP_TAY|" & Text.From(r[DongNguon]), MaKheUoc = Ma, NgayGiaoDich = Ngay,
            GocDaTra = try So(r[GocDaTra]) otherwise 0, LaiPhiDaTra = try So(r[LaiPhiDaTra]) otherwise 0,
            NoiDung = GhiChu, TenFileNguon = "TRA_NO_NHAP_TAY", DongNguon = r[DongNguon], SoTienGhi = null,
            LyDoLoai = LyDo,
            ChiTiet = if Thieu then "Cần đủ ngày giao dịch, mã khế ước và số tiền" else GhiChu]),
    // ---- Sao kê
    SK = Table.SelectRows(SaoKe, each List.Contains({"GOC", "LAI_PHI", "CHO_PHAN_LOAI"}, [PhanLoai])),
    Sao = List.Transform(Table.ToRecords(SK), (r) => [
        Nguon = "SAO_KE", IDGiaoDich = r[IDGiaoDich], MaKheUoc = r[MaKheUoc], NgayGiaoDich = r[NgayGiaoDich],
        GocDaTra = r[GocDaTra], LaiPhiDaTra = r[LaiPhiDaTra], NoiDung = r[NoiDung], TenFileNguon = r[TenFileNguon],
        DongNguon = r[DongNguon], SoTienGhi = r[GhiNo],
        LyDoLoai = if r[PhanLoai] = "CHO_PHAN_LOAI" then "Sao kê liên quan vay chưa tách được gốc/lãi (nhập tay nếu cần)"
            else if not List.Contains(MaHopLe, r[MaKheUoc]) then "Mã khế ước không có trong danh mục (không tính)"
            else null,
        ChiTiet = r[NoiDung]]),
    SaoTinh = Table.FromRecords(List.Select(Sao, each [LyDoLoai] = null), Kieu),
    KhoaSaoKe = Table.Buffer(Table.Group(SaoTinh, {"MaKheUoc", "NgayGiaoDich"}, {
        {"SkGoc", each List.Sum([GocDaTra]), type number}, {"SkLai", each List.Sum([LaiPhiDaTra]), type number}})),
    ChongTrung = List.Transform(Tay, (r) =>
        if r[LyDoLoai] <> null then r
        else
            let T = Table.SelectRows(KhoaSaoKe, each [MaKheUoc] = r[MaKheUoc] and [NgayGiaoDich] = r[NgayGiaoDich])
            in if Table.IsEmpty(T) then r
               else Record.TransformFields(r, {
                   {"LyDoLoai", each "Nhập tay trùng ngày với sao kê – đã bỏ dòng nhập tay"},
                   {"ChiTiet", each "Sao kê cùng ngày: gốc " & Number.ToText(T{0}[SkGoc], "N0", "en-US")
                       & ", lãi/phí " & Number.ToText(T{0}[SkLai], "N0", "en-US")}})),
    KetQua = Table.Buffer(Table.FromRecords(Sao & ChongTrung, Kieu))
in
    KetQua
