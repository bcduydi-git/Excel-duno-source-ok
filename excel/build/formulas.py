"""Công thức các sheet báo cáo (Excel 365, dynamic array). Viết dạng dễ đọc rồi chuyển sang dạng lưu trong file
(thêm tiền tố _xlfn./_xlpm.). Không dùng LAMBDA/MAP và tham chiếu A1# để LibreOffice kiểm thử được."""
import re

FUNCS = {
    "LET": "_xlfn.LET", "FILTER": "_xlfn._xlws.FILTER", "SORT": "_xlfn._xlws.SORT", "SORTBY": "_xlfn.SORTBY",
    "UNIQUE": "_xlfn.UNIQUE", "XLOOKUP": "_xlfn.XLOOKUP", "HSTACK": "_xlfn.HSTACK", "VSTACK": "_xlfn.VSTACK",
    "CHOOSECOLS": "_xlfn.CHOOSECOLS", "SEQUENCE": "_xlfn.SEQUENCE", "MAXIFS": "_xlfn.MAXIFS", "MINIFS": "_xlfn.MINIFS",
    "TAKE": "_xlfn.TAKE", "DROP": "_xlfn.DROP",
}


def to_file(formula):
    """Chuyển công thức dễ đọc sang dạng lưu trong file: tiền tố hàm mới và tên biến LET."""
    let_vars = set()
    for m in re.finditer(r"LET\(", formula):
        pass
    # Tên biến LET: mọi định danh đứng ngay sau "LET(" hoặc sau dấu phẩy ở cấp 1 của LET, theo sau là dấu phẩy.
    let_vars = set(re.findall(r"(?:LET\(|,)\s*([a-z][a-z_]*)\s*,", formula))
    out, i, n = [], 0, len(formula)
    while i < n:
        c = formula[i]
        if c == '"':
            j = formula.index('"', i + 1)
            while j + 1 < n and formula[j + 1] == '"':
                j = formula.index('"', j + 2)
            out.append(formula[i:j + 1]); i = j + 1; continue
        if c == "[":
            depth, j = 0, i
            while True:
                if formula[j] == "[": depth += 1
                elif formula[j] == "]":
                    depth -= 1
                    if depth == 0: break
                j += 1
            out.append(formula[i:j + 1]); i = j + 1; continue
        m = re.match(r"[A-Za-z_][A-Za-z0-9_.]*", formula[i:])
        if m and (i == 0 or not (formula[i - 1].isalnum() or formula[i - 1] in "$_.")):
            w = m.group(0)
            nxt = formula[i + len(w):i + len(w) + 1]
            if nxt == "(" and w in FUNCS:
                out.append(FUNCS[w])
            elif w in let_vars and nxt != "[":
                out.append("_xlpm." + w)
            else:
                out.append(w)
            i += len(w); continue
        out.append(c); i += 1
    return "".join(out)


NGUONG = 'XLOOKUP("NguongLechLai",tblThamSo[ThamSo],tblThamSo[GiaTri],0)'

# Dư nợ từng khế ước tại ngày `d` (biểu thức dùng lại): gốc ban đầu − gốc thực trả (từ ngày làm gốc) đến ngày d.
def du_no_expr(ma, gbd, bd, d):
    return (f"IF(({bd}>0)*({bd}<={d}),{gbd}-SUMIFS(KQ_GiaoDich[GocDaTra],KQ_GiaoDich[MaKheUoc],{ma},"
            f"KQ_GiaoDich[NgayGiaoDich],\"<=\"&{d},KQ_GiaoDich[TinhDuNo],1),0)")


# ------------------------------------------------------------------ 4. Lịch trả nợ
# Trạng thái từng kỳ tính đến "Đến ngày" (quy tắc Power BI + các chỉnh sửa đã thống nhất):
#  gốc: phân bổ gốc thực trả luỹ kế theo thứ tự kỳ; lãi/phí: tổng giao dịch đã ghép vào kỳ (MaKy);
#  kỳ có gốc trả hết → xong lãi; đã có giao dịch lãi mà thiếu ≤ ngưỡng → đã trả;
#  khoản đã tất toán → mọi kỳ xong; kỳ trước lần trả lãi đầu tiên có dữ liệu → coi như đã trả.
def lich_tra_no(r0):
    return f"""=LET(
tu,$B$3, den,$B$4, pn,$B$5, nhf,$B$6, ttf,$B$7, moc,MIN(den,TODAY()), ng,{NGUONG},
sel,(KQ_LichTraNo[NgayDenHan]>=tu)*(KQ_LichTraNo[NgayDenHan]<=den)*((pn="")+(KQ_LichTraNo[MaCongTy]=pn)>0)*((nhf="")+(KQ_LichTraNo[MaNganHang]=nhf)>0),
ngay,FILTER(KQ_LichTraNo[NgayDenHan],sel,0),
cty,FILTER(KQ_LichTraNo[MaCongTy],sel,""),
nhh,FILTER(KQ_LichTraNo[MaNganHang],sel,""),
ma,FILTER(KQ_LichTraNo[MaKheUoc],sel,""),
kyn,FILTER(KQ_LichTraNo[KyTraNo],sel,0),
maky,FILTER(KQ_LichTraNo[MaKy],sel,""),
goc,FILTER(KQ_LichTraNo[GocPhaiTra],sel,0),
lai,FILTER(KQ_LichTraNo[LaiPhiDuKien],sel,0),
luyke,FILTER(KQ_LichTraNo[GocLuyKe],sel,0),
truoc,FILTER(KQ_LichTraNo[TruocDuLieu],sel,0),
nguon,FILTER(KQ_LichTraNo[NguonLich],sel,""),
gtra,SUMIFS(KQ_GiaoDich[GocDaTra],KQ_GiaoDich[MaKheUoc],ma,KQ_GiaoDich[NgayGiaoDich],"<="&den,KQ_GiaoDich[TinhDuNo],1),
ttoan,XLOOKUP(ma,KQ_KheUoc[MaKheUoc],KQ_KheUoc[NgayTatToan],0),
tat,(ttoan>0)*(ttoan<=den),
gcon,IF(tat,0,IF(luyke-gtra<=0,0,IF(luyke-gtra>goc,goc,luyke-gtra))),
ltra,IF(truoc=1,0,SUMIFS(KQ_GiaoDich[LaiPhiDaTra],KQ_GiaoDich[MaKy],maky,KQ_GiaoDich[NgayGiaoDich],"<="&den)),
sogd,COUNTIFS(KQ_GiaoDich[MaKy],maky,KQ_GiaoDich[NgayGiaoDich],"<="&den,KQ_GiaoDich[LaiPhiDaTra],">0"),
lcon,IF(tat+truoc>0,0,IF((goc>0)*(gcon=0),0,IF(lai<=0,0,IF(sogd=0,lai,IF(lai-ltra>ng,lai-ltra,0))))),
st,IF(tat*(ngay>ttoan),"Đã tất toán",IF(gcon+lcon=0,"Đã trả",IF(ngay<=moc,"Chưa trả","Chưa đến hạn"))),
gchu,IF(truoc=1,"Trước lần trả lãi đầu tiên có dữ liệu",IF(tat,"Tất toán "&TEXT(ttoan,"dd/mm/yyyy"),"")),
ngtra,MAXIFS(KQ_GiaoDich[NgayGiaoDich],KQ_GiaoDich[MaKy],maky,KQ_GiaoDich[NgayGiaoDich],"<="&den),
bang,HSTACK(ngay,cty,nhh,ma,kyn,goc,lai,goc-gcon,ltra,IF(ngtra=0,"",ngtra),gcon,lcon,st,gchu,nguon),
keep,IF(ttf="",ngay=ngay,st=ttf),
kq,FILTER(bang,keep,""),
tong,HSTACK("TỔNG: "&ROWS(kq)&" kỳ","","","","",SUM(CHOOSECOLS(kq,6)),SUM(CHOOSECOLS(kq,7)),SUM(CHOOSECOLS(kq,8)),SUM(CHOOSECOLS(kq,9)),"",SUM(CHOOSECOLS(kq,11)),SUM(CHOOSECOLS(kq,12)),"","",""),
IF(SUM(sel)*SUM(--keep)=0,"Không có kỳ nào thỏa điều kiện lọc",VSTACK(tong,kq)))"""


LICH_COLS = ["Ngày đến hạn", "Pháp nhân", "Ngân hàng", "Mã khế ước", "Kỳ", "Gốc phải trả", "Lãi/phí dự kiến",
             "Gốc đã trả", "Lãi/phí đã trả", "Ngày trả lãi gần nhất", "Gốc còn phải trả", "Lãi/phí còn phải trả",
             "Trạng thái", "Ghi chú", "Nguồn lịch"]


# ------------------------------------------------------------------ 3. Danh sách khế ước
def khe_uoc():
    du_no = du_no_expr("ma", "gbd", "bd", "d")
    return f"""=LET(
d,$B$3, moc,MIN(d,TODAY()), pn,$B$4, nhf,$B$5, ttf,$B$6,
sel,((pn="")+(KQ_KheUoc[MaCongTy]=pn)>0)*((nhf="")+(KQ_KheUoc[MaNganHang]=nhf)>0),
ma,FILTER(KQ_KheUoc[MaKheUoc],sel,""),
cty,FILTER(KQ_KheUoc[MaCongTy],sel,""),
nhh,FILTER(KQ_KheUoc[MaNganHang],sel,""),
hd,FILTER(KQ_KheUoc[MaHDTD],sel,""),
nhan,FILTER(KQ_KheUoc[NgayNhanNo],sel,0),
dao,FILTER(KQ_KheUoc[NgayDaoHan],sel,0),
gbd,FILTER(KQ_KheUoc[GocBanDau],sel,0),
bd,FILTER(KQ_KheUoc[NgayBatDau],sel,0),
ls,FILTER(KQ_KheUoc[LaiSuatNam],sel,0),
gtra,SUMIFS(KQ_GiaoDich[GocDaTra],KQ_GiaoDich[MaKheUoc],ma,KQ_GiaoDich[NgayGiaoDich],"<="&d,KQ_GiaoDich[TinhDuNo],1),
duno,{du_no},
ltra,SUMIFS(KQ_GiaoDich[LaiPhiDaTra],KQ_GiaoDich[MaKheUoc],ma,KQ_GiaoDich[NgayGiaoDich],"<="&d),
con,duno>0.5,
ketiep,MINIFS(KQ_LichTraNo[NgayDenHan],KQ_LichTraNo[MaKheUoc],ma,KQ_LichTraNo[NgayDenHan],">"&moc),
songay,IF(con*(dao>0)*(dao>=moc),dao-moc,""),
quahan,IF(con*(dao>0)*(dao<moc),moc-dao,""),
st,IF(dao=0,"Thiếu ngày đáo hạn",IF(con=0,IF((nhan>0)*(nhan>moc),"Chưa giải ngân","Đã tất toán"),IF(dao<moc,"Quá hạn",IF(dao-moc<=7,"Đến hạn ≤ 7 ngày",IF(dao-moc<=30,"Đến hạn ≤ 30 ngày","Trong hạn"))))),
bang,HSTACK(cty,nhh,ma,IF(hd="","",hd),IF(nhan=0,"",nhan),IF(dao=0,"",dao),gbd,IF(ls=0,"",ls),gtra,duno,ltra,IF(con*(ketiep>0),ketiep,""),songay,quahan,st),
keep,IF(ttf="",ma=ma,IF(ttf="Còn dư nợ",con=1,st=ttf)),
kq,SORTBY(FILTER(bang,keep,""),FILTER(cty,keep,""),1,FILTER(nhh,keep,""),1,FILTER(dao,keep,0),1),
tong,HSTACK("TỔNG: "&ROWS(kq)&" khế ước","","","","","",SUM(CHOOSECOLS(kq,7)),"",SUM(CHOOSECOLS(kq,9)),SUM(CHOOSECOLS(kq,10)),SUM(CHOOSECOLS(kq,11)),"","","",""),
IF(SUM(sel)*SUM(--keep)=0,"Không có khế ước nào thỏa điều kiện lọc",VSTACK(tong,kq)))"""


KHE_UOC_COLS = ["Pháp nhân", "Ngân hàng", "Mã khế ước", "Mã HĐTD", "Ngày nhận nợ", "Ngày đáo hạn",
                "Số vay / dư nợ đầu theo dõi", "Lãi suất", "Gốc đã trả", "Dư nợ", "Lãi/phí đã trả",
                "Kỳ trả kế tiếp", "Số ngày đến đáo hạn", "Số ngày quá hạn", "Trạng thái"]


# ------------------------------------------------------------------ 2. HĐTD & hạn mức
def hop_dong():
    du_no = du_no_expr("kma", "kgbd", "kbd", "d")
    return f"""=LET(
d,$B$3, pn,$B$4, nhf,$B$5,
sel,((pn="")+(tblHopDong[MaCongTy]=pn)>0)*((nhf="")+(tblHopDong[MaNganHang]=nhf)>0)*(tblHopDong[MaHDTD]<>""),
cty,FILTER(tblHopDong[MaCongTy],sel,"~"),
hd,FILTER(tblHopDong[MaHDTD],sel,"~"),
nhh,FILTER(tblHopDong[MaNganHang],sel,""),
hl,FILTER(tblHopDong[NgayHieuLuc],sel,0),
het,FILTER(tblHopDong[NgayHetHan],sel,0),
hm,FILTER(tblHopDong[HanMucTong],sel,0),
md,FILTER(tblHopDong[MucDichVay],sel,""),
ts,FILTER(tblHopDong[TSDB],sel,""),
kma,KQ_KheUoc[MaKheUoc], kgbd,KQ_KheUoc[GocBanDau], kbd,KQ_KheUoc[NgayBatDau],
kduno,{du_no},
khop,(cty=TRANSPOSE(KQ_KheUoc[MaCongTy]))*(hd=TRANSPOSE(KQ_KheUoc[MaHDTD])),
duno,MMULT(khop,kduno),
soku,MMULT(khop,--(kduno>0.5)),
st,IF(hl>d,"Chưa hiệu lực",IF((het>0)*(het<d),"Hết hạn","Đang hiệu lực")),
hieuluc,st="Đang hiệu lực",
bang,HSTACK(cty,hd,nhh,IF(hl=0,"",hl),IF(het=0,"",het),st,hm,duno,IF(hieuluc,hm-duno,""),IF(hieuluc*(hm>0),duno/IF(hm=0,1,hm),""),soku,IF(md="","",md),IF(ts="","",ts)),
hmhl,SUM(hm*hieuluc), dnhl,SUM(duno*hieuluc),
tong,HSTACK("TỔNG ("&ROWS(bang)&" HĐ) – hạn mức đang hiệu lực","","","","","",hmhl,dnhl,hmhl-dnhl,IF(hmhl>0,dnhl/hmhl,""),SUM(soku),"",""),
IF(SUM(sel)=0,"Không có hợp đồng nào thỏa điều kiện lọc",VSTACK(tong,bang)))"""


HOP_DONG_COLS = ["Pháp nhân", "Mã HĐTD", "Ngân hàng", "Ngày hiệu lực", "Ngày hết hạn", "Trạng thái", "Hạn mức",
                 "Dư nợ thuộc HĐ", "Hạn mức chưa sử dụng", "Tỷ lệ sử dụng", "Số KU còn dư nợ", "Mục đích vay", "TSĐB"]


def ngoai_han_muc():
    """Dư nợ không thuộc hợp đồng đang hiệu lực (HĐ hết hạn, khoản không có mã HĐ như Chailease, mã HĐ không khớp)."""
    du_no = du_no_expr("kma", "kgbd", "kbd", "d")
    return f"""=LET(
d,$B$3, pn,$B$4, nhf,$B$5,
kma,KQ_KheUoc[MaKheUoc], kgbd,KQ_KheUoc[GocBanDau], kbd,KQ_KheUoc[NgayBatDau],
loc,((pn="")+(KQ_KheUoc[MaCongTy]=pn)>0)*((nhf="")+(KQ_KheUoc[MaNganHang]=nhf)>0),
kduno,{du_no}*loc,
hlhd,(tblHopDong[NgayHieuLuc]<=d)*((tblHopDong[NgayHetHan]="")+(tblHopDong[NgayHetHan]>=d)>0)*(tblHopDong[MaHDTD]<>""),
trong,MMULT(--((KQ_KheUoc[MaCongTy]&"|"&KQ_KheUoc[MaHDTD])=TRANSPOSE(tblHopDong[MaCongTy]&"|"&tblHopDong[MaHDTD])),hlhd),
SUM(kduno*(trong=0)))"""


# ------------------------------------------------------------------ 1. Tổng quan
def tong_quan():
    dn_dau = du_no_expr("ma", "gbd", "bd", "(tu-1)")
    dn_cuoi = du_no_expr("ma", "gbd", "bd", "den")
    return f"""=LET(
tu,$B$3, den,$B$4, pn,$B$5, nhf,$B$6, moc,MIN(den,TODAY()), ng,{NGUONG},
sel,((pn="")+(KQ_KheUoc[MaCongTy]=pn)>0)*((nhf="")+(KQ_KheUoc[MaNganHang]=nhf)>0),
ma,FILTER(KQ_KheUoc[MaKheUoc],sel,"~"),
cty,FILTER(KQ_KheUoc[MaCongTy],sel,"~"),
nhh,FILTER(KQ_KheUoc[MaNganHang],sel,"~"),
gbd,FILTER(KQ_KheUoc[GocBanDau],sel,0),
bd,FILTER(KQ_KheUoc[NgayBatDau],sel,0),
nhan,FILTER(KQ_KheUoc[NgayNhanNo],sel,0),
vay,FILTER(KQ_KheUoc[SoTienVay],sel,0),
dau,{dn_dau},
cuoi,{dn_cuoi},
giaingan,IF((nhan>=tu)*(nhan<=den),vay,0),
tragoc,SUMIFS(KQ_GiaoDich[GocDaTra],KQ_GiaoDich[MaKheUoc],ma,KQ_GiaoDich[NgayGiaoDich],">="&tu,KQ_GiaoDich[NgayGiaoDich],"<="&den,KQ_GiaoDich[TinhDuNo],1),
unq,UNIQUE(HSTACK(cty,nhh)),
nhom,SORTBY(unq,CHOOSECOLS(unq,1),1,CHOOSECOLS(unq,2),1),
kc,CHOOSECOLS(nhom,1), kn,CHOOSECOLS(nhom,2),
mk,(kc=TRANSPOSE(cty))*(kn=TRANSPOSE(nhh)),
lsel,(KQ_LichTraNo[NgayDenHan]>=tu)*(KQ_LichTraNo[NgayDenHan]<=den)*((pn="")+(KQ_LichTraNo[MaCongTy]=pn)>0)*((nhf="")+(KQ_LichTraNo[MaNganHang]=nhf)>0),
lcty,FILTER(KQ_LichTraNo[MaCongTy],lsel,"~"),
lnh,FILTER(KQ_LichTraNo[MaNganHang],lsel,"~"),
lma,FILTER(KQ_LichTraNo[MaKheUoc],lsel,""),
lky,FILTER(KQ_LichTraNo[MaKy],lsel,""),
lngay,FILTER(KQ_LichTraNo[NgayDenHan],lsel,0),
lgoc,FILTER(KQ_LichTraNo[GocPhaiTra],lsel,0),
llai,FILTER(KQ_LichTraNo[LaiPhiDuKien],lsel,0),
lluy,FILTER(KQ_LichTraNo[GocLuyKe],lsel,0),
ltruoc,FILTER(KQ_LichTraNo[TruocDuLieu],lsel,0),
lgtra,SUMIFS(KQ_GiaoDich[GocDaTra],KQ_GiaoDich[MaKheUoc],lma,KQ_GiaoDich[NgayGiaoDich],"<="&den,KQ_GiaoDich[TinhDuNo],1),
lttoan,XLOOKUP(lma,KQ_KheUoc[MaKheUoc],KQ_KheUoc[NgayTatToan],0),
ltat,(lttoan>0)*(lttoan<=den),
lgcon,IF(ltat,0,IF(lluy-lgtra<=0,0,IF(lluy-lgtra>lgoc,lgoc,lluy-lgtra))),
lltra,SUMIFS(KQ_GiaoDich[LaiPhiDaTra],KQ_GiaoDich[MaKy],lky,KQ_GiaoDich[NgayGiaoDich],"<="&den),
lsogd,COUNTIFS(KQ_GiaoDich[MaKy],lky,KQ_GiaoDich[NgayGiaoDich],"<="&den,KQ_GiaoDich[LaiPhiDaTra],">0"),
llcon,IF(ltat+ltruoc>0,0,IF((lgoc>0)*(lgcon=0),0,IF(llai<=0,0,IF(lsogd=0,llai,IF(llai-lltra>ng,llai-lltra,0))))),
ml,(kc=TRANSPOSE(lcty))*(kn=TRANSPOSE(lnh)),
laitra,SUMIFS(KQ_GiaoDich[LaiPhiDaTra],KQ_GiaoDich[MaCongTy],kc,KQ_GiaoDich[MaNganHang],kn,KQ_GiaoDich[NgayGiaoDich],">="&tu,KQ_GiaoDich[NgayGiaoDich],"<="&den),
bang,HSTACK(kc,kn,MMULT(mk,--(cuoi>0.5)),MMULT(mk,dau),MMULT(mk,giaingan),MMULT(mk,tragoc),MMULT(mk,cuoi),MMULT(ml,lgoc),MMULT(ml,llai),laitra,MMULT(ml,lgcon),MMULT(ml,llcon)),
tong,HSTACK("TỔNG","",SUM(CHOOSECOLS(bang,3)),SUM(CHOOSECOLS(bang,4)),SUM(CHOOSECOLS(bang,5)),SUM(CHOOSECOLS(bang,6)),SUM(CHOOSECOLS(bang,7)),SUM(CHOOSECOLS(bang,8)),SUM(CHOOSECOLS(bang,9)),SUM(CHOOSECOLS(bang,10)),SUM(CHOOSECOLS(bang,11)),SUM(CHOOSECOLS(bang,12))),
IF(SUM(sel)=0,"Không có khế ước nào thỏa điều kiện lọc",VSTACK(tong,bang)))"""


TONG_QUAN_COLS = ["Pháp nhân", "Ngân hàng", "Số KU còn dư nợ", "Dư nợ đầu kỳ", "Giải ngân trong kỳ", "Trả gốc trong kỳ",
                  "Dư nợ cuối kỳ", "Gốc đến hạn trong kỳ", "Lãi/phí dự kiến đến hạn", "Lãi/phí đã trả trong kỳ",
                  "Gốc đến hạn còn phải trả", "Lãi/phí đến hạn còn phải trả"]


def tong_quan_han_muc():
    """[Hạn mức đang hiệu lực, dư nợ trong hạn mức hiệu lực, dư nợ quá hạn] tại Đến ngày (một hàng 3 ô)."""
    du_no = du_no_expr("kma", "kgbd", "kbd", "den")
    return f"""=LET(
den,$B$4, pn,$B$5, nhf,$B$6, moc,MIN(den,TODAY()),
kma,KQ_KheUoc[MaKheUoc], kgbd,KQ_KheUoc[GocBanDau], kbd,KQ_KheUoc[NgayBatDau],
loc,((pn="")+(KQ_KheUoc[MaCongTy]=pn)>0)*((nhf="")+(KQ_KheUoc[MaNganHang]=nhf)>0),
kduno,{du_no}*loc,
hlhd,(tblHopDong[NgayHieuLuc]<=den)*((tblHopDong[NgayHetHan]="")+(tblHopDong[NgayHetHan]>=den)>0)*(tblHopDong[MaHDTD]<>"")*((pn="")+(tblHopDong[MaCongTy]=pn)>0)*((nhf="")+(tblHopDong[MaNganHang]=nhf)>0),
trong,MMULT(--((KQ_KheUoc[MaCongTy]&"|"&KQ_KheUoc[MaHDTD])=TRANSPOSE(tblHopDong[MaCongTy]&"|"&tblHopDong[MaHDTD])),hlhd),
quahan,(KQ_KheUoc[NgayDaoHan]>0)*(KQ_KheUoc[NgayDaoHan]<moc)*(kduno>0.5),
HSTACK(SUM(tblHopDong[HanMucTong]*hlhd),SUM(kduno*(trong>0)),SUM(kduno*quahan)))"""
