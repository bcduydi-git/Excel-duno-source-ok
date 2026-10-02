// Giao dịch trả nợ được tính (nhập tay + sao kê, đã chống trùng). Truy nguồn bằng TenFileNguon + DongNguon.
// TinhDuNo = 0: giao dịch trước mốc theo dõi của khoản chỉ có lịch dài hạn (không trừ vào dư nợ).
// MaKy: kỳ được ghép cho giao dịch lãi/phí — theo NGÀY, không theo số tiền. Một "lần trả" = mọi giao dịch
//   lãi/phí của cùng khế ước trong cùng ngày; mỗi lần trả thuộc đúng một kỳ:
//   - lần trả đầu tiên có dữ liệu: kỳ đến hạn gần nhất trước ngày trả (kỳ đầu nếu trả trước hạn);
//   - lần sau: kỳ đã đến hạn chưa có lần trả nào, cũ nhất; nếu không còn kỳ như vậy thì kỳ gần ngày trả
//     nhất giữa kỳ vừa ghép (trả bổ sung) và kỳ kế tiếp (trả trước hạn).
let
    GD = Table.SelectRows(GiaoDichNguon, each [LyDoLoai] = null),
    KU = Table.SelectColumns(DanhMucKheUoc, {"MaKheUoc", "MaCongTy", "MaNganHang", "MaHDTD", "NgayBatDauTheoDoi"}),
    Ghep = Table.ExpandTableColumn(Table.NestedJoin(GD, {"MaKheUoc"}, KU, {"MaKheUoc"}, "KU", JoinKind.LeftOuter),
        "KU", {"MaCongTy", "MaNganHang", "MaHDTD", "NgayBatDauTheoDoi"}),
    ThemTinh = Table.AddColumn(Ghep, "TinhDuNo", each
        if [NgayBatDauTheoDoi] <> null and [NgayGiaoDich] < [NgayBatDauTheoDoi] then 0 else 1, Int64.Type),
    Lich = Table.SelectRows(LichKhung, each [Loi] = null),
    NhomLich = Table.Group(Lich, {"MaKheUoc"}, {{"Ky", each Table.ToRecords(Table.Sort(
        Table.SelectColumns(_, {"MaKy", "KyTraNo", "NgayDenHan"}),
        {{"NgayDenHan", Order.Ascending}, {"KyTraNo", Order.Ascending}}))}}),
    TraCuuLich = Record.FromList(NhomLich[Ky], NhomLich[MaKheUoc]),
    LaiPhi = Table.SelectRows(ThemTinh, each [LaiPhiDaTra] > 0 and [TinhDuNo] = 1),
    NgayTheoMa = Table.Group(LaiPhi, {"MaKheUoc"}, {{"Ngay", each List.Sort(List.Distinct([NgayGiaoDich]))}}),
    GhepKy = (Ma as text, CacNgay as list) as list =>
        let
            Ky = Record.FieldOrDefault(TraCuuLich, Ma, {}),
            N = List.Count(Ky),
            Cach = (d as date, i as nullable number) as number => Number.Abs(Duration.Days(d - Ky{i}[NgayDenHan])),
            KQ = List.Accumulate(CacNgay, [i = null, cuoi = null, out = {}], (s, d) =>
                let
                    j = List.Count(List.Select(Ky, each [NgayDenHan] <= d)) - 1,
                    c = if s[i] = null then List.Max({j, 0})
                        else if s[i] <= j then s[i]
                        else if s[i] < N and Cach(d, s[i]) < Cach(d, s[cuoi]) then s[i]
                        else s[cuoi]
                in
                    [i = List.Max({if s[i] = null then 0 else s[i], c + 1}), cuoi = c,
                     out = s[out] & {[MaKheUoc = Ma, NgayGiaoDich = d, MaKy = Ky{c}[MaKy], NgayDenHanKy = Ky{c}[NgayDenHan]]}])
        in
            if N = 0 then {} else KQ[out],
    BangKy = Table.FromRecords(List.Combine(List.Transform(Table.ToRecords(NgayTheoMa), (r) => GhepKy(r[MaKheUoc], r[Ngay]))),
        type table [MaKheUoc = text, NgayGiaoDich = date, MaKy = text, NgayDenHanKy = date]),
    NoiKy = Table.ExpandTableColumn(Table.NestedJoin(ThemTinh, {"MaKheUoc", "NgayGiaoDich"}, BangKy,
        {"MaKheUoc", "NgayGiaoDich"}, "K", JoinKind.LeftOuter), "K", {"MaKy", "NgayDenHanKy"}),
    MoKy = Table.ReplaceValue(NoiKy, each [MaKy], each if [LaiPhiDaTra] > 0 and [TinhDuNo] = 1 then [MaKy] else null,
        Replacer.ReplaceValue, {"MaKy"}),
    MoKy2 = Table.ReplaceValue(MoKy, each [NgayDenHanKy], each if [MaKy] = null then null else [NgayDenHanKy],
        Replacer.ReplaceValue, {"NgayDenHanKy"}),
    Chon = Table.SelectColumns(MoKy2, {"Nguon", "IDGiaoDich", "MaCongTy", "MaNganHang", "MaHDTD", "MaKheUoc",
        "NgayGiaoDich", "GocDaTra", "LaiPhiDaTra", "TinhDuNo", "MaKy", "NgayDenHanKy", "NoiDung", "TenFileNguon", "DongNguon"}),
    Kieu = Table.TransformColumnTypes(Chon, {{"NgayGiaoDich", type date}, {"GocDaTra", type number},
        {"LaiPhiDaTra", type number}, {"DongNguon", Int64.Type}, {"MaKy", type text}, {"NgayDenHanKy", type date}}),
    KetQua = Table.Buffer(Table.Sort(Kieu, {{"NgayGiaoDich", Order.Ascending}, {"MaKheUoc", Order.Ascending},
        {"IDGiaoDich", Order.Ascending}}))
in
    KetQua
