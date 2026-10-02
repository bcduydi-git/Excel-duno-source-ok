// Danh mục khế ước: KHE_UOC + các khoản chỉ có lịch ở LICH_DAI_HAN (như tblKheUoc của Power BI cũ).
// Khoản chỉ có lịch: dư nợ ban đầu = DuNoTheoLichDauKy của dòng lịch đầu tiên từ mốc theo dõi.
let
    Moc = ThamSo[MocTheoDoiLichDaiHan],
    ChuanChu = (Bang as table, Cot as list) as table =>
        Table.TransformColumns(Bang, List.Transform(Cot, (c) => {c, each
            if _ = null then null else let t = Text.Trim(Text.From(_)) in if t = "" then null else t,
            type nullable text})),
    ChuanNgay = (Bang as table, Cot as list) as table =>
        Table.TransformColumns(Bang, List.Transform(Cot, (c) => {c, each
            if _ = null or (_ is text and Text.Trim(_) = "") then null else Date.From(_, "vi-VN"),
            type nullable date})),
    Nguon = Excel.CurrentWorkbook(){[Name = "tblKheUoc"]}[Content],
    Chu = ChuanChu(Nguon, {"MaCongTy", "MaNganHang", "MaKheUoc", "MaHDTD", "PhuongThucTraGoc",
        "TanSuatTraGoc", "TanSuatTraLai", "GhiChu"}),
    CoMa = Table.SelectRows(Chu, each [MaKheUoc] <> null),
    Ngay = ChuanNgay(CoMa, {"NgayNhanNo", "NgayDaoHan", "NgayTraGocDauTien", "NgayTraLaiDauTien"}),
    So = Table.TransformColumnTypes(Ngay, {{"SoTienVay", type nullable number},
        {"LaiSuatNam", type nullable number}, {"CoSoTinhLai", Int64.Type}}, "vi-VN"),
    KheUoc = Table.AddColumn(Table.AddColumn(So, "DuNoBanDau", each null, type nullable number),
        "NgayBatDauTheoDoi", each null, type nullable date),
    Lich = Excel.CurrentWorkbook(){[Name = "tblLichDaiHan"]}[Content],
    LichChu = ChuanChu(Lich, {"MaCongTy", "MaNganHang", "MaKheUoc"}),
    LichNgay = ChuanNgay(Table.SelectRows(LichChu, each [MaKheUoc] <> null), {"NgayDenHan"}),
    MaDaCo = List.Buffer(KheUoc[MaKheUoc]),
    CanBoSung = Table.SelectRows(LichNgay, each [NgayDenHan] <> null and [NgayDenHan] >= Moc
        and not List.Contains(MaDaCo, [MaKheUoc])),
    Gom = Table.Group(CanBoSung, {"MaCongTy", "MaNganHang", "MaKheUoc"}, {
        {"DuNoBanDau", each
            let d = Table.Sort(_, {{"NgayDenHan", Order.Ascending}}){0}[DuNoTheoLichDauKy]
            in if d = null then null else Number.From(d), type nullable number},
        {"NgayDaoHan", each List.Max([NgayDenHan]), type date}}),
    BoSung = Table.AddColumn(Table.AddColumn(Gom, "NgayBatDauTheoDoi", each Moc, type date),
        "PhuongThucTraGoc", each "THEO_LICH", type text),
    Gop = Table.Combine({KheUoc, BoSung}),
    ThemGoc = Table.AddColumn(Gop, "GocBanDau", each
        if [NgayBatDauTheoDoi] <> null then [DuNoBanDau] else [SoTienVay], type nullable number),
    ThemNgay = Table.AddColumn(ThemGoc, "NgayBatDau", each
        if [NgayBatDauTheoDoi] <> null then [NgayBatDauTheoDoi] else [NgayNhanNo], type nullable date),
    KetQua = Table.Buffer(ThemNgay)
in
    KetQua
