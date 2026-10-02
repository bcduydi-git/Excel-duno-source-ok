// Lịch trả nợ hoàn chỉnh: khung lịch + lãi/phí dự kiến.
// - Lịch ngân hàng (LICH_DAI_HAN): lãi/phí = TongPhaiTra - GocPhaiTra.
// - TruocDuLieu = 1: kỳ đến hạn trước kỳ của lần trả lãi đầu tiên có dữ liệu → coi như đã trả (như Power BI cũ:
//   không coi thiếu giao dịch lịch sử là chưa trả).
// - Kỳ tự sinh: lãi dự kiến tính trên dư nợ thực tế từng ngày trong kỳ (đã trừ gốc trả trước hạn),
//   = Σ dư nợ × số ngày × lãi suất / cơ sở tính lãi. Chỉ là ước tính; số thực trả lấy ở KQ_GiaoDich.
let
    KU = Table.Buffer(Table.SelectColumns(DanhMucKheUoc, {"MaKheUoc", "SoTienVay", "LaiSuatNam", "CoSoTinhLai"})),
    TraCuuKU = Record.FromList(Table.ToRecords(KU), KU[MaKheUoc]),
    Tra = Table.Group(Table.SelectRows(KQ_GiaoDich, each [TinhDuNo] = 1 and [GocDaTra] <> 0), {"MaKheUoc"},
        {{"DS", each List.Buffer(Table.ToRecords(Table.SelectColumns(_, {"NgayGiaoDich", "GocDaTra"})))}}),
    TraCuuTra = Record.FromList(Tra[DS], Tra[MaKheUoc]),
    LaiDuKien = (r as record) as nullable number =>
        if r[NguonLich] <> "TU_SINH" then r[LaiPhiNhap]
        else
            let
                K = Record.Field(TraCuuKU, r[MaKheUoc]),
                Tu = r[TuNgayTinhLai],
                Den = r[NgayDenHan],
                DS = Record.FieldOrDefault(TraCuuTra, r[MaKheUoc], {}),
                Moc = List.Sort(List.Distinct({Tu, Den}
                    & List.Select(List.Transform(DS, each [NgayGiaoDich]), each _ > Tu and _ < Den))),
                Tong = List.Sum({0} & List.Transform(List.Positions(List.RemoveLastN(Moc, 1)), (i) =>
                    let
                        A = Moc{i},
                        DuNo = K[SoTienVay] - List.Sum({0} & List.Transform(List.Select(DS, each [NgayGiaoDich] <= A), each [GocDaTra]))
                    in
                        List.Max({0, DuNo}) * Duration.Days(Moc{i + 1} - A)))
            in
                if K[LaiSuatNam] = null or not List.Contains({360, 365}, K[CoSoTinhLai]) then null
                else Number.Round(Tong * K[LaiSuatNam] / K[CoSoTinhLai], 0, RoundingMode.AwayFromZero),
    KyDau = Table.Group(Table.SelectRows(KQ_GiaoDich, each [NgayDenHanKy] <> null), {"MaKheUoc"},
        {{"Ngay", each List.Min([NgayDenHanKy]), type date}}),
    TraCuuKyDau = Record.FromList(KyDau[Ngay], KyDau[MaKheUoc]),
    Lich = Table.SelectRows(LichKhung, each [Loi] = null),
    ThemLai = Table.AddColumn(Lich, "LaiPhiDuKien", each LaiDuKien(_), type nullable number),
    ThemSoNgay = Table.AddColumn(ThemLai, "SoNgayTinhLai", each
        if [TuNgayTinhLai] = null then null else Duration.Days([NgayDenHan] - [TuNgayTinhLai]), type nullable number),
    ThemTong = Table.AddColumn(ThemSoNgay, "TongPhaiTra", each
        if [LaiPhiDuKien] = null then null else [GocPhaiTra] + [LaiPhiDuKien], type nullable number),
    ThemTruoc = Table.AddColumn(ThemTong, "TruocDuLieu", each
        let d = Record.FieldOrDefault(TraCuuKyDau, [MaKheUoc], null) in if d <> null and [NgayDenHan] < d then 1 else 0, Int64.Type),
    Chon = Table.SelectColumns(ThemTruoc, {"MaCongTy", "MaNganHang", "MaKheUoc", "MaKy", "KyTraNo", "TuNgayTinhLai",
        "NgayDenHan", "SoNgayTinhLai", "GocPhaiTra", "LaiPhiDuKien", "TongPhaiTra", "GocLuyKe", "TruocDuLieu", "NguonLich"}),
    Kieu = Table.TransformColumnTypes(Chon, {{"KyTraNo", Int64.Type}, {"TuNgayTinhLai", type date},
        {"NgayDenHan", type date}, {"GocPhaiTra", type number}, {"GocLuyKe", type number}}),
    KetQua = Table.Sort(Kieu, {{"NgayDenHan", Order.Ascending}, {"MaCongTy", Order.Ascending},
        {"MaNganHang", Order.Ascending}, {"MaKheUoc", Order.Ascending}})
in
    KetQua
