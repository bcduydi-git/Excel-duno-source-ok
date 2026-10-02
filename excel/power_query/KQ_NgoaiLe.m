// Những gì cần người dùng xem: chỉ các trường hợp không xử lý tự động được hoặc dữ liệu bất thường.
let
    Chu = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    Kieu = type table [Loai = text, MaKheUoc = nullable text, NgayGiaoDich = nullable date, SoTien = nullable number,
        ChiTiet = nullable text, Nguon = nullable text, TenFileNguon = nullable text, DongNguon = nullable Int64.Type],
    Dong = (Loai as text, r as record, ChiTiet as nullable text) as record => [
        Loai = Loai, MaKheUoc = r[MaKheUoc]?, NgayGiaoDich = r[NgayGiaoDich]?,
        SoTien = if r[SoTienGhi]? <> null then r[SoTienGhi] else (r[GocDaTra]? ?? 0) + (r[LaiPhiDaTra]? ?? 0),
        ChiTiet = ChiTiet, Nguon = r[Nguon]?, TenFileNguon = r[TenFileNguon]?, DongNguon = r[DongNguon]?],
    // 1. Khoản không sinh được lịch
    Lich = List.Transform(Table.ToRecords(Table.SelectRows(LichKhung, each [Loi] <> null)), (r) =>
        [Loai = "Không sinh được lịch trả nợ", MaKheUoc = r[MaKheUoc], NgayGiaoDich = null, SoTien = null,
         ChiTiet = r[Loi], Nguon = null, TenFileNguon = null, DongNguon = null]),
    // 2. Giao dịch không được tính (thiếu thông tin, sai mã, trùng sao kê, sao kê chưa phân loại)
    Loai = List.Transform(Table.ToRecords(Table.SelectRows(GiaoDichNguon, each [LyDoLoai] <> null)),
        (r) => Dong(r[LyDoLoai], r, r[ChiTiet])),
    // 3. Giao dịch lãi không ghép được kỳ
    KhongKy = List.Transform(Table.ToRecords(Table.SelectRows(KQ_GiaoDich, each [LaiPhiDaTra] > 0 and [TinhDuNo] = 1 and [MaKy] = null)),
        (r) => Dong("Không xác định được kỳ trả lãi", r, "Khế ước chưa có lịch trả nợ")),
    // 4. Giao dịch trước ngày nhận nợ (vẫn tính như Power BI cũ)
    NhanNo = Record.FromList(KQ_KheUoc[NgayNhanNo], KQ_KheUoc[MaKheUoc]),
    TruocNhanNo = List.Transform(List.Select(Table.ToRecords(KQ_GiaoDich), each
            let n = Record.Field(NhanNo, [MaKheUoc]) in n <> null and [NgayGiaoDich] < n),
        (r) => Dong("Ngày trả trước ngày nhận nợ (vẫn tính)", r,
            "Nhận nợ " & Date.ToText(Record.Field(NhanNo, r[MaKheUoc]), [Format = "dd/MM/yyyy"]))),
    // 5. Gốc đã trả vượt số vay
    TongGoc = Table.Group(Table.SelectRows(KQ_GiaoDich, each [TinhDuNo] = 1), {"MaKheUoc"}, {{"Tong", each List.Sum([GocDaTra]), type number}}),
    VuotGoc = List.Transform(List.Select(Table.ToRecords(Table.NestedJoin(TongGoc, {"MaKheUoc"},
            Table.SelectColumns(KQ_KheUoc, {"MaKheUoc", "GocBanDau"}), {"MaKheUoc"}, "K", JoinKind.Inner)),
            each [K]{0}[GocBanDau] <> null and [Tong] > [K]{0}[GocBanDau] + 0.5),
        (r) => [Loai = "Gốc đã trả vượt số vay", MaKheUoc = r[MaKheUoc], NgayGiaoDich = null,
            SoTien = r[Tong] - r[K]{0}[GocBanDau],
            ChiTiet = "Đã trả " & Number.ToText(r[Tong], "N0", "en-US") & " / vay " & Number.ToText(r[K]{0}[GocBanDau], "N0", "en-US"),
            Nguon = null, TenFileNguon = null, DongNguon = null]),
    // 6. Thư mục sao kê
    TM = Table.SelectRows(Excel.CurrentWorkbook(){[Name = "tblThuMucSaoKe"]}[Content], each Chu([MaNganHang]) <> ""),
    ThuMuc = List.Combine(List.Transform(Table.ToRecords(TM), (r) =>
        let
            NH = Chu(r[MaNganHang]),
            SoFile = try Table.RowCount(Table.SelectRows(Folder.Files(Chu(r[ThuMuc])), each
                List.Contains({".xlsx", ".xls"}, Text.Lower([Extension])) and not Text.StartsWith([Name], "~$"))) otherwise null,
            Loi = if not List.Contains({"VPBank", "NABank"}, NH) then "Chưa có bộ đọc sao kê tự động cho ngân hàng này (nhập tay)"
                else if SoFile = null then "Không mở được thư mục sao kê"
                else if SoFile = 0 then "Thư mục sao kê chưa có file"
                else null
        in
            if Loi = null then {} else {[Loai = Loi, MaKheUoc = null, NgayGiaoDich = null, SoTien = null,
                ChiTiet = Chu(r[MaCongTy]) & " / " & NH & ": " & Chu(r[ThuMuc]), Nguon = "SAO_KE", TenFileNguon = null, DongNguon = null]})),
    Bang = Table.FromRecords(ThuMuc & Lich & Loai & KhongKy & TruocNhanNo & VuotGoc, Kieu),
    KetQua = Table.Sort(Bang, {{"Loai", Order.Ascending}, {"MaKheUoc", Order.Ascending},
        {"NgayGiaoDich", Order.Ascending}, {"DongNguon", Order.Ascending}})
in
    KetQua
