// Đọc một file sao kê NABank (giữ nguyên cách đọc của Power BI cũ: q_SaoKeNABank).
// Trả về: DongNguon, SoGiaoDich, NgayGiaoDich, GhiNo, NoiDung, MaKheUoc, PhanLoai.
(NoiDungFile as binary, TenFile as text) as table =>
let
    Chu = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    SoTien = (x as any) as number =>
        if Chu(x) = "" then 0 else if x is number then x else Number.FromText(Chu(x), "en-US"),
    Ngay = (x as any) as nullable date =>
        if Chu(x) = "" then null
        else if x is text then Date.FromText(Chu(x), [Format = "dd/MM/yyyy", Culture = "vi-VN"])
        else Date.From(x),
    Workbook = Excel.Workbook(NoiDungFile, [UseHeaders = false, DelayTypes = true, InferSheetDimensions = true]),
    Sheet = Workbook{[Item = "Sao kê", Kind = "Sheet"]}[Data],
    DanhSo = Table.AddIndexColumn(Sheet, "DongNguon", 1, 1, Int64.Type),
    LocGiaoDich = Table.SelectRows(DanhSo, each
        let STT = Chu([Column1]) in STT <> "" and Text.Select(STT, {"0".."9"}) = STT),
    Doi = Table.FromRecords(List.Transform(Table.ToRecords(LocGiaoDich), (r) => [
            DongNguon = r[DongNguon],
            SoGiaoDich = Chu(r[Column4]),
            NgayGiaoDich = Ngay(r[Column2]),
            GhiNo = SoTien(r[Column6]),
            NoiDung = Chu(r[Column5])]),
        type table [DongNguon = Int64.Type, SoGiaoDich = text, NgayGiaoDich = date, GhiNo = number, NoiDung = text]),
    // Nội dung thu nợ NABank có hai dạng:
    //   "Thu von khe uoc: <mã>(<ref>)"  - không dấu, có dấu hai chấm
    //   "Thu vốn khế ước <mã> (<ref>)"  - có dấu, không dấu hai chấm
    TienToGoc = {"THU VON KHE UOC", "THU VỐN KHẾ ƯỚC"},
    TienToLai = {"THU LAI KHE UOC", "THU LÃI KHẾ ƯỚC"},
    LoaiThu = (ND as text) as text =>
        if List.MatchesAny(TienToGoc, each Text.StartsWith(ND, _)) then "GOC"
        else if List.MatchesAny(TienToLai, each Text.StartsWith(ND, _)) then "LAI_PHI"
        else "",
    TachMa = (NoiDung as text) as nullable text =>
        let
            ND = Text.Upper(NoiDung),
            Loai = LoaiThu(ND),
            CacTienTo = if Loai = "GOC" then TienToGoc else if Loai = "LAI_PHI" then TienToLai else {},
            TienTo = List.First(List.Select(CacTienTo, each Text.StartsWith(ND, _)), null),
            PhanSau = if TienTo = null then "" else Text.Trim(Text.Range(NoiDung, Text.Length(TienTo))),
            BoHaiCham = if Text.StartsWith(PhanSau, ":") then Text.Trim(Text.Range(PhanSau, 1)) else PhanSau,
            Ma = Text.Trim(Text.BeforeDelimiter(BoHaiCham & "(", "("))
        in
            if Ma = "" then null else Ma,
    ThemMa = Table.AddColumn(Doi, "MaKheUoc", each TachMa([NoiDung]), type nullable text),
    ThemLoai = Table.AddColumn(ThemMa, "PhanLoai", each
        let
            ND = Text.Upper([NoiDung]),
            CoMa = [MaKheUoc] <> null,
            GocLaiGop = Text.Contains(ND, "GOC VA LAI") or Text.Contains(ND, "GỐC VÀ LÃI"),
            PhiLC = Text.Contains(ND, "LC") and (Text.Contains(ND, "PHI") or Text.Contains(ND, "PHÍ"))
        in
            if [GhiNo] <= 0 then "KHAC"
            else if LoaiThu(ND) = "GOC" and CoMa then "GOC"
            else if LoaiThu(ND) = "LAI_PHI" and CoMa then "LAI_PHI"
            else if GocLaiGop or PhiLC or Text.Contains(ND, "KHE UOC") or Text.Contains(ND, "KHẾ ƯỚC")
                or Text.Contains(ND, "THU NO") or Text.Contains(ND, "THU NỢ") then "CHO_PHAN_LOAI"
            else "KHAC",
        type text)
in
    ThemLoai
