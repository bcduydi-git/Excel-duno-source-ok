// Gộp mọi file sao kê trong các thư mục khai báo ở CAU_HINH (bảng tblThuMucSaoKe).
// Chỉ VPBank và NABank được đọc tự động (như Power BI cũ); ngân hàng khác nhập tay.
// Thư mục không mở được hoặc ngân hàng chưa có bộ đọc: bỏ qua ở đây, báo ở KQ_NgoaiLe.
let
    Chu = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    KieuRong = type table [
        IDGiaoDich = text, MaCongTy = text, NganHangSaoKe = text, MaKheUoc = nullable text,
        NgayGiaoDich = date, GhiNo = number, GocDaTra = number, LaiPhiDaTra = number,
        PhanLoai = text, NoiDung = text, SoGiaoDich = text, TenFileNguon = text, DongNguon = Int64.Type],
    CauHinh = Excel.CurrentWorkbook(){[Name = "tblThuMucSaoKe"]}[Content],
    CoThuMuc = Table.SelectRows(CauHinh, each Chu([ThuMuc]) <> "" and List.Contains({"VPBank", "NABank"}, Chu([MaNganHang]))),
    DocThuMuc = (r as record) as list =>
        let
            MaCty = Chu(r[MaCongTy]),
            NH = Chu(r[MaNganHang]),
            DS = try Table.Buffer(Folder.Files(Chu(r[ThuMuc]))) otherwise null,
            Loc = if DS = null then {} else Table.ToRecords(Table.SelectRows(DS, each
                List.Contains({".xlsx", ".xls"}, Text.Lower([Extension]))
                and not Text.StartsWith([Name], "~$")
                and [Attributes]?[Hidden]? <> true)),
            Doc = List.Transform(Loc, (f) =>
                let
                    Bang = if NH = "VPBank" then fnSaoKeVPBank(f[Content], f[Name]) else fnSaoKeNABank(f[Content], f[Name]),
                    ThemCty = Table.AddColumn(Bang, "MaCongTy", each MaCty, type text),
                    ThemNH = Table.AddColumn(ThemCty, "NganHangSaoKe", each NH, type text),
                    ThemFile = Table.AddColumn(ThemNH, "TenFileNguon", each f[Name], type text),
                    ThemID = Table.AddColumn(ThemFile, "IDGiaoDich", each MaCty & "|" & NH & "|" & f[Name] & "|" & Text.From([DongNguon]), type text)
                in
                    ThemID)
        in
            Doc,
    CacBang = List.Combine(List.Transform(Table.ToRecords(CoThuMuc), DocThuMuc)),
    Gop = Table.Combine(CacBang & {#table(KieuRong, {})}),
    ThemGoc = Table.AddColumn(Gop, "GocDaTra2", each if [PhanLoai] = "GOC" then [GhiNo] else 0, type number),
    ThemLai = Table.AddColumn(ThemGoc, "LaiPhiDaTra2", each if [PhanLoai] = "LAI_PHI" then [GhiNo] else 0, type number),
    BoCu = Table.RemoveColumns(ThemLai, {"GocDaTra", "LaiPhiDaTra"}, MissingField.Ignore),
    DoiTen = Table.RenameColumns(BoCu, {{"GocDaTra2", "GocDaTra"}, {"LaiPhiDaTra2", "LaiPhiDaTra"}}),
    KetQua = Table.SelectColumns(DoiTen, Table.ColumnNames(#table(KieuRong, {})))
in
    KetQua
