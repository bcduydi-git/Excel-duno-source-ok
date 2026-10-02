// Khung lịch trả nợ (chưa tính lãi dự kiến cho kỳ tự sinh).
// - Khoản có lịch ở LICH_DAI_HAN: dùng trực tiếp lịch đó (từ mốc theo dõi nếu có).
// - Khoản trả gốc cuối kỳ: tự sinh kỳ lãi theo NgayTraLaiDauTien, hoặc ngày trả lãi cố định
//   của ngân hàng (DANH_MUC, cột NgayTraLai); kỳ cuối luôn là ngày đáo hạn (như Power BI cũ).
// Khoản không sinh được lịch: một dòng có cột Loi (đưa sang KQ_NgoaiLe), không làm hỏng Refresh.
let
    Chu = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    KU = DanhMucKheUoc,
    NH = Table.SelectRows(Excel.CurrentWorkbook(){[Name = "tblNganHang"]}[Content], each Chu([MaNganHang]) <> ""),
    NgayCoDinh = Record.FromList(
        List.Transform(NH[NgayTraLai], each if Chu(_) = "" then null else Number.From(_)),
        List.Transform(NH[MaNganHang], Chu)),
    LichNguon = Excel.CurrentWorkbook(){[Name = "tblLichDaiHan"]}[Content],
    LichNhap = Table.Buffer(Table.FromRecords(List.Transform(
        Table.ToRecords(Table.SelectRows(LichNguon, each Chu([MaKheUoc]) <> "" and Chu([NgayDenHan]) <> "")),
        (r) => [
            MaKheUoc = Chu(r[MaKheUoc]),
            KyTraNo = Number.From(r[KyTraNo]),
            NgayDenHan = Date.From(r[NgayDenHan], "vi-VN"),
            GocPhaiTra = if Chu(r[GocPhaiTra]) = "" then null else Number.From(r[GocPhaiTra]),
            TongPhaiTra = if Chu(r[TongPhaiTra]) = "" then null else Number.From(r[TongPhaiTra])]))),
    CoLich = List.Buffer(List.Distinct(LichNhap[MaKheUoc])),
    TaoLich = (r as record) as list =>
        if List.Contains(CoLich, r[MaKheUoc]) then
            let
                BatDau = if r[NgayBatDauTheoDoi] = null then #date(1900, 1, 1) else r[NgayBatDauTheoDoi],
                DS = Table.ToRecords(Table.Sort(
                    Table.SelectRows(LichNhap, each [MaKheUoc] = r[MaKheUoc] and [NgayDenHan] >= BatDau),
                    {{"NgayDenHan", Order.Ascending}, {"KyTraNo", Order.Ascending}}))
            in
                List.Transform(List.Positions(DS), (i) => [
                    KyTraNo = DS{i}[KyTraNo],
                    TuNgayTinhLai = if i = 0 then null else DS{i - 1}[NgayDenHan],
                    NgayDenHan = DS{i}[NgayDenHan],
                    GocPhaiTra = if DS{i}[GocPhaiTra] = null then 0 else DS{i}[GocPhaiTra],
                    LaiPhiNhap = if DS{i}[TongPhaiTra] = null or DS{i}[GocPhaiTra] = null then null
                        else DS{i}[TongPhaiTra] - DS{i}[GocPhaiTra],
                    NguonLich = "LICH_NGAN_HANG",
                    Loi = null])
        else
            let
                NhanNo = r[NgayNhanNo],
                DaoHan = r[NgayDaoHan],
                SoTien = r[SoTienVay],
                TanSuat = r[TanSuatTraLai],
                Buoc = if TanSuat = "THANG" then 1 else if TanSuat = "QUY" then 3
                    else if TanSuat = "6_THANG" then 6 else if TanSuat = "CUOI_KY" then 0 else null,
                CoDinh = Record.FieldOrDefault(NgayCoDinh, Chu(r[MaNganHang]), null),
                NgayDau =
                    if r[NgayTraLaiDauTien] <> null then r[NgayTraLaiDauTien]
                    else if Buoc = null or Buoc = 0 or CoDinh = null or NhanNo = null then null
                    else let d = #date(Date.Year(NhanNo), Date.Month(NhanNo), CoDinh)
                         in if d > NhanNo then d else Date.AddMonths(d, 1),
                Loi =
                    if r[PhuongThucTraGoc] <> "CUOI_KY" then "Chưa có lịch ở LICH_DAI_HAN và không tự sinh được với phương thức trả gốc này"
                    else if NhanNo = null or DaoHan = null then "Thiếu ngày nhận nợ hoặc ngày đáo hạn"
                    else if DaoHan <= NhanNo then "Ngày đáo hạn phải sau ngày nhận nợ"
                    else if SoTien = null or SoTien <= 0 then "Thiếu hoặc sai số tiền vay"
                    else if Buoc = null then "Tần suất trả lãi không hợp lệ"
                    else if Buoc > 0 and NgayDau = null then "Cần ngày trả lãi đầu tiên (KHE_UOC) hoặc ngày trả lãi cố định của ngân hàng (DANH_MUC)"
                    else if Buoc > 0 and NgayDau <= NhanNo then "Ngày trả lãi đầu tiên phải sau ngày nhận nợ"
                    else null,
                CacNgay =
                    (if Buoc = 0 then {}
                     else List.Generate(() => 0, (i) => Date.AddMonths(NgayDau, i * Buoc) < DaoHan,
                         (i) => i + 1, (i) => Date.AddMonths(NgayDau, i * Buoc))) & {DaoHan}
            in
                if Loi <> null then
                    {[KyTraNo = null, TuNgayTinhLai = null, NgayDenHan = null, GocPhaiTra = null,
                      LaiPhiNhap = null, NguonLich = null, Loi = Loi]}
                else
                    List.Transform(List.Positions(CacNgay), (i) => [
                        KyTraNo = i + 1,
                        TuNgayTinhLai = if i = 0 then NhanNo else CacNgay{i - 1},
                        NgayDenHan = CacNgay{i},
                        GocPhaiTra = if CacNgay{i} = DaoHan then SoTien else 0,
                        LaiPhiNhap = null,
                        NguonLich = "TU_SINH",
                        Loi = null]),
    DongLich = List.Combine(List.Transform(Table.ToRecords(KU), (r) =>
        List.Transform(TaoLich(r), (k) =>
            Record.Combine({[MaCongTy = r[MaCongTy], MaNganHang = r[MaNganHang], MaKheUoc = r[MaKheUoc]], k})))),
    Bang = Table.FromRecords(DongLich, type table [
        MaCongTy = nullable text, MaNganHang = nullable text, MaKheUoc = text, KyTraNo = nullable Int64.Type,
        TuNgayTinhLai = nullable date, NgayDenHan = nullable date, GocPhaiTra = nullable number,
        LaiPhiNhap = nullable number, NguonLich = nullable text, Loi = nullable text]),
    CoLoi = Table.SelectRows(Bang, each [Loi] <> null),
    KhongLoi = Table.SelectRows(Bang, each [Loi] = null),
    // Gốc phải trả luỹ kế đến hết từng kỳ (để phân bổ gốc đã trả theo thứ tự kỳ).
    LuyKe = Table.Combine(Table.Group(KhongLoi, {"MaKheUoc"}, {{"B", each
        let
            S = Table.Sort(_, {{"NgayDenHan", Order.Ascending}, {"KyTraNo", Order.Ascending}}),
            G = List.Buffer(S[GocPhaiTra]),
            L = List.Skip(List.Accumulate(G, {0}, (a, x) => a & {List.Last(a) + x}), 1)
        in
            Table.FromColumns(Table.ToColumns(S) & {L}, Table.ColumnNames(S) & {"GocLuyKe"})}})[B]
        & {Table.AddColumn(Table.FirstN(KhongLoi, 0), "GocLuyKe", each null)}),
    ThemMaKy = Table.AddColumn(LuyKe, "MaKy", each [MaKheUoc] & "|" & Text.From([KyTraNo]), type text),
    KetQua = Table.Buffer(Table.Combine({ThemMaKy, CoLoi}))
in
    KetQua
