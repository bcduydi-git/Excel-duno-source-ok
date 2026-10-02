// Danh mục khế ước dùng cho báo cáo. GocBanDau/NgayBatDau: số và ngày làm gốc tính dư nợ
// (khoản thường: số tiền vay, ngày nhận nợ; khoản chỉ có lịch dài hạn: dư nợ đầu, mốc theo dõi).
// NgayTatToan: ngày gốc thực trả luỹ kế đủ GocBanDau (theo giao dịch thực tế, không suy đoán).
let
    CT = Excel.CurrentWorkbook(){[Name = "tblCongTy"]}[Content],
    TenCty = Record.FromList(CT[TenCongTy], List.Transform(CT[MaCongTy], each if _ = null then "" else Text.Trim(Text.From(_)))),
    Tra = Table.Group(Table.SelectRows(KQ_GiaoDich, each [TinhDuNo] = 1 and [GocDaTra] <> 0), {"MaKheUoc"},
        {{"DS", each Table.ToRecords(Table.Sort(Table.SelectColumns(_, {"NgayGiaoDich", "GocDaTra"}),
            {{"NgayGiaoDich", Order.Ascending}}))}}),
    TraCuuTra = Record.FromList(Tra[DS], Tra[MaKheUoc]),
    NgayTatToan = (Ma as text, Goc as nullable number) as nullable date =>
        let
            DS = Record.FieldOrDefault(TraCuuTra, Ma, {}),
            LuyKe = List.Accumulate(DS, [Tong = 0, Ngay = null], (s, x) =>
                let T = s[Tong] + x[GocDaTra]
                in [Tong = T, Ngay = if s[Ngay] = null and Goc <> null and T >= Goc - 0.5 then x[NgayGiaoDich] else s[Ngay]])
        in
            LuyKe[Ngay],
    ThemTen = Table.AddColumn(DanhMucKheUoc, "TenCongTy", each Record.FieldOrDefault(TenCty, [MaCongTy], null), type nullable text),
    ThemTatToan = Table.AddColumn(ThemTen, "NgayTatToan", each NgayTatToan([MaKheUoc], [GocBanDau]), type nullable date),
    KetQua = Table.SelectColumns(ThemTatToan, {"MaCongTy", "TenCongTy", "MaNganHang", "MaKheUoc", "MaHDTD",
        "NgayNhanNo", "NgayDaoHan", "SoTienVay", "LaiSuatNam", "PhuongThucTraGoc", "TanSuatTraLai",
        "NgayTraLaiDauTien", "CoSoTinhLai", "DuNoBanDau", "NgayBatDauTheoDoi", "GocBanDau", "NgayBatDau",
        "NgayTatToan", "GhiChu"})
in
    KetQua
