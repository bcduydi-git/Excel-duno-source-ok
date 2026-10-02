// Đọc một file sao kê VPBank (giữ nguyên cách đọc của Power BI cũ: q_SaoKeVPBank).
// Trả về: DongNguon, SoGiaoDich, NgayGiaoDich, GhiNo, NoiDung, MaKheUoc, PhanLoai.
(NoiDungFile as binary, TenFile as text) as table =>
let
    Chu = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    Chuan = (x as any) as text =>
        Text.Upper(Text.Trim(Text.Replace(Text.Replace(Chu(x), "#(cr)", " "), "#(lf)", " "))),
    SoTien = (x as any) as number =>
        let
            S = Text.Replace(Text.Replace(Chu(x), " ", ""), Character.FromNumber(160), ""),
            DauCham = Text.PositionOf(S, ".", Occurrence.Last),
            DauPhay = Text.PositionOf(S, ",", Occurrence.Last),
            VanHoa =
                if DauCham >= 0 and DauPhay >= 0 then (if DauPhay > DauCham then "vi-VN" else "en-US")
                else if DauCham >= 0 then
                    (if List.Count(Text.Split(S, ".")) > 2 then "vi-VN"
                     else if Text.Length(Text.AfterDelimiter(S, ".")) = 3 then "vi-VN" else "en-US")
                else if DauPhay >= 0 then
                    (if List.Count(Text.Split(S, ",")) > 2 then "en-US"
                     else if Text.Length(Text.AfterDelimiter(S, ",")) = 3 then "en-US" else "vi-VN")
                else "en-US"
        in
            if x is number then x else if S = "" then 0 else Number.FromText(S, VanHoa),
    Ngay = (x as any) as nullable date =>
        if Chu(x) = "" then null
        else if x is text then Date.FromText(Text.Start(Chu(x), 10), [Format = "dd/MM/yyyy", Culture = "vi-VN"])
        else Date.From(x),
    ViTriTieuDe = (Bang as table) as number =>
        let
            CacDong = Table.ToRows(Table.FirstN(Bang, 40)),
            DanhDau = List.Transform(CacDong, (Dong) =>
                let ND = List.Transform(Dong, each Chuan(_))
                in List.AnyTrue(List.Transform(ND, each Text.Contains(_, "NỘI DUNG GIAO DỊCH")))
                    and List.AnyTrue(List.Transform(ND, each Text.Contains(_, "SỐ TIỀN GHI NỢ"))))
        in
            List.PositionOf(DanhDau, true),
    Workbook = Excel.Workbook(NoiDungFile, null, true),
    CacBang = List.Select(Table.ToRecords(Workbook), each Value.Is(Record.FieldOrDefault(_, "Data", null), type table)),
    UngVien = List.Transform(CacBang, (R) => [
        Ten = Chu(Record.FieldOrDefault(R, "Item", Record.FieldOrDefault(R, "Name", ""))),
        Bang = Record.Field(R, "Data"),
        ViTri = ViTriTieuDe(Record.Field(R, "Data"))]),
    BangHopLe = List.Select(UngVien, each [ViTri] >= 0),
    UuTien = List.Select(BangHopLe, each Chuan([Ten]) = "TỔNG HỢP"),
    ChonBang =
        if List.Count(UuTien) > 0 then UuTien{0}
        else if List.Count(BangHopLe) > 0 then BangHopLe{0}
        else error Error.Record("KhongTimThayBangSaoKe", "Không tìm thấy tiêu đề giao dịch VPBank trong file.", [TenFile = TenFile]),
    Bang = ChonBang[Bang],
    ViTri = ChonBang[ViTri],
    TenCot = Table.ColumnNames(Bang),
    TieuDe = List.Transform(Table.ToRows(Table.Range(Bang, ViTri, 1)){0}, each Chuan(_)),
    TimCot = (Nhan as text) as text =>
        let p = List.PositionOf(List.Transform(TieuDe, each Text.StartsWith(_, Chuan(Nhan))), true)
        in if p < 0 then error Error.Record("ThieuCot", "Không tìm thấy cột " & Nhan, [TenFile = TenFile]) else TenCot{p},
    CotSTT = TimCot("STT"),
    CotSoGD = TimCot("Số bút toán"),
    CotNgay = TimCot("Ngày giá trị"),
    CotNo = TimCot("Số tiền ghi nợ"),
    CotND = TimCot("Nội dung giao dịch"),
    CotGio = TimCot("Thời gian giao dịch"),
    DanhSo = Table.AddIndexColumn(Bang, "DongNguon", 1, 1, Int64.Type),
    LocGiaoDich = Table.SelectRows(Table.Skip(DanhSo, ViTri + 1), each
        let STT = Chu(Record.Field(_, CotSTT)) in STT <> "" and Text.Select(STT, {"0".."9"}) = STT),
    Doi = Table.FromRecords(List.Transform(Table.ToRecords(LocGiaoDich), (r) =>
        let Gio = Record.Field(r, CotGio)
        in [
            DongNguon = r[DongNguon],
            SoGiaoDich = Chu(Record.Field(r, CotSoGD)),
            NgayGiaoDich = if Chu(Gio) <> "" then Ngay(Gio) else Ngay(Record.Field(r, CotNgay)),
            GhiNo = SoTien(Record.Field(r, CotNo)),
            NoiDung = Chu(Record.Field(r, CotND))]),
        type table [DongNguon = Int64.Type, SoGiaoDich = text, NgayGiaoDich = date, GhiNo = number, NoiDung = text]),
    ThemMa = Table.AddColumn(Doi, "MaKheUoc", each
        let
            BoHauTo = Text.BeforeDelimiter(Text.Upper([SoGiaoDich]) & ";", ";"),
            Ma = if Text.StartsWith(BoHauTo, "PDLD") then Text.Range(BoHauTo, 2) else BoHauTo
        in
            if Text.StartsWith(Ma, "LD") and Text.Length(Ma) = 12
                and Text.Select(Text.Range(Ma, 2), {"0".."9"}) = Text.Range(Ma, 2)
            then Ma else null,
        type nullable text),
    ThemLoai = Table.AddColumn(ThemMa, "PhanLoai", each
        let
            ND = Chuan([NoiDung]),
            CoMa = [MaKheUoc] <> null,
            LaGoc = Text.StartsWith(ND, "THANH TOAN GOC") or Text.StartsWith(ND, "THANH TOAN NO GOC"),
            LaLai = Text.StartsWith(ND, "THANH TOAN LAI"),
            GopGocLai = Text.Contains(ND, "GOC VA LAI") or Text.Contains(ND, "GỐC VÀ LÃI")
        in
            if [GhiNo] <= 0 then "KHAC"
            else if GopGocLai then "CHO_PHAN_LOAI"
            else if LaGoc and CoMa then "GOC"
            else if LaLai and CoMa then "LAI_PHI"
            else if CoMa or LaGoc or LaLai or Text.Contains(ND, "THU NO")
                or Text.Contains(ND, "THU GOC") or Text.Contains(ND, "THU LAI") then "CHO_PHAN_LOAI"
            else "KHAC",
        type text)
in
    ThemLoai
