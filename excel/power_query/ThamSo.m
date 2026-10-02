// Tham số dùng chung, nhập ở sheet CAU_HINH (bảng tblThamSo).
let
    Bang = Excel.CurrentWorkbook(){[Name = "tblThamSo"]}[Content],
    CoTen = Table.SelectRows(Bang, each [ThamSo] <> null and Text.Trim(Text.From([ThamSo])) <> ""),
    GiaTri = Record.FromList(CoTen[GiaTri], List.Transform(CoTen[ThamSo], each Text.Trim(Text.From(_)))),
    KetQua = [
        MocTheoDoiLichDaiHan = Date.From(GiaTri[MocTheoDoiLichDaiHan], "vi-VN"),
        NguongLechLai = Number.From(GiaTri[NguongLechLai])
    ]
in
    KetQua
