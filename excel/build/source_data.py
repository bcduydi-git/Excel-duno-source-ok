"""Đọc các bảng nguồn từ file Excel nhập liệu (giống Excel.CurrentWorkbook trong Power Query).

Dữ liệu thật KHÔNG nằm trong repo (repo public). Đường dẫn truyền qua tham số.
"""
import datetime as dt
import openpyxl
from openpyxl.utils import range_boundaries


def _cell(v):
    if isinstance(v, str):
        v = v.strip()
        return v if v != "" else None
    return v


def read_tables(path):
    """Trả về {tên bảng: list[dict]} cho mọi Excel Table trong workbook (giá trị đã tính)."""
    wb_f = openpyxl.load_workbook(path)                    # để lấy định nghĩa bảng
    wb_v = openpyxl.load_workbook(path, data_only=True)    # để lấy giá trị đã tính
    out = {}
    for ws in wb_f.worksheets:
        wv = wb_v[ws.title]
        for t in ws.tables.values():
            c1, r1, c2, r2 = range_boundaries(t.ref)
            rows = list(wv.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True))
            hdr = [str(h).strip() for h in rows[0]]
            data = []
            for i, r in enumerate(rows[1:]):
                d = {h: _cell(v) for h, v in zip(hdr, r)}
                d["__row"] = r1 + 1 + i
                data.append(d)
            out[t.name] = {"sheet": ws.title, "ref": t.ref, "cols": hdr, "rows": data}
    return out


def as_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        return dt.datetime.strptime(v.strip(), "%d/%m/%Y").date()
    if isinstance(v, (int, float)):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=v)).date()
    raise ValueError(v)


def as_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    return float(s) if s else 0.0
