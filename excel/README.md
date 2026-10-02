# Báo cáo dư nợ trên Excel (thay Power BI)

File dùng: `Bao_cao_du_no.xlsx`, là bản sao file nhập liệu `Du_lieu_dau_vao_du_no.xlsx` có thêm phần xử lý và báo cáo.
**Repo này public nên không chứa file Excel hay dữ liệu thật.** Chỉ có mã nguồn (query, script dựng file).

## Cấu trúc file

| Sheet | Vai trò | Cập nhật |
|---|---|---|
| `1_TONG_QUAN`, `2_HDTD_HAN_MUC`, `3_KHE_UOC`, `4_LICH_TRA_NO` | 4 báo cáo, viết bằng công thức Excel 365 | Đổi ô vàng là kết quả cập nhật ngay |
| `DANH_MUC`, `HOP_DONG`, `KHE_UOC`, `LICH_DAI_HAN`, `TRA_NO_NHAP_TAY` | Nhập liệu (giữ nguyên như cũ) | Nhập tay |
| `CAU_HINH` | Tham số và thư mục sao kê | Nhập tay |
| `DL_KheUoc`, `DL_LichTraNo`, `DL_GiaoDich`, `NGOAI_LE` | Kết quả Power Query, xem chi tiết hoặc lọc | Data > Refresh All |

Thay đổi so với file nguồn (chỉ ở những chỗ cần thiết):
- `DANH_MUC`: bảng ngân hàng thêm cột `NgayTraLai` (25 / 20 / 15). Trước đây các ngày này viết cứng trong query Power BI.
- `KHE_UOC`: thấu chi HDBank có ngày trả lãi đầu tiên 15/10/2025, thay cho ngoại lệ viết cứng trong query cũ.
- Thêm sheet `CAU_HINH`.

## Quy tắc tính (theo Power BI cũ, cộng các chỉnh sửa đã thống nhất)

- **Dư nợ** = số vay − gốc thực trả đến ngày báo cáo. Khoản chỉ có lịch ở `LICH_DAI_HAN` (Chailease) lấy dư nợ đầu theo lịch tại mốc theo dõi 01/09/2026, rồi trừ gốc trả từ mốc này.
- **Số thực trả** lấy từ `TRA_NO_NHAP_TAY` và sao kê VPBank, NABank (đọc tự động từ thư mục ở `CAU_HINH`). Cột lãi/phí gồm mọi khoản ngoài gốc.
- **Chống tính trùng:** dòng nhập tay cùng mã khế ước và cùng ngày với giao dịch sao kê đã đọc thì bị bỏ, và được liệt kê ở `NGOAI_LE`.
- **Lịch trả nợ:** khoản có lịch ở `LICH_DAI_HAN` thì dùng thẳng lịch đó. Khoản trả gốc cuối kỳ thì tự sinh kỳ lãi theo ngày trả lãi đầu tiên hoặc ngày cố định của ngân hàng; kỳ cuối là ngày đáo hạn.
- **Lãi dự kiến** của kỳ tự sinh tính theo dư nợ thực tế từng ngày, tức đã trừ gốc trả trước hạn. Đây chỉ là ước tính; không dùng lãi dự kiến thay cho số thực trả.
- **Ghép giao dịch lãi vào kỳ** theo ngày, không theo số tiền. Các giao dịch lãi/phí của cùng khế ước trong cùng ngày được tính là một lần trả. Mỗi lần trả thuộc đúng một kỳ:
  - Lần trả đầu tiên có dữ liệu ghép vào kỳ đến hạn gần nhất trước ngày trả.
  - Lần sau ghép vào kỳ đã đến hạn chưa có lần trả nào, cũ nhất.
  - Nếu không còn kỳ như vậy, ghép vào kỳ gần ngày trả nhất.
- **Trạng thái kỳ**, tính đến "Đến ngày":
  - Gốc: gốc thực trả luỹ kế được phân bổ theo thứ tự kỳ.
  - Kỳ có gốc đã trả hết được coi là đã xong lãi.
  - Kỳ đã có giao dịch lãi mà thiếu so với dự kiến không quá `NguongLechLai` (5 triệu) vẫn là "Đã trả".
  - Khoản đã tất toán (gốc thực trả đủ) thì mọi kỳ đến hạn trước ngày tất toán là "Đã trả"; các kỳ sau đó là "Đã tất toán".
  - Kỳ trước lần trả lãi đầu tiên có dữ liệu được coi là đã trả, vì thiếu lịch sử không có nghĩa là chưa trả.

## Kiểm chứng

- Công thức 4 báo cáo được LibreOffice 26.2 tính lại với 5 bộ tham số (ngày, pháp nhân, ngân hàng, trạng thái, kể cả trường hợp không có kết quả). Kết quả khớp từng ô với bản tính Python `build/expected.py`.
- `build/compare_pbi.py` mô phỏng các measure DAX của Power BI cũ để đối chiếu. Đây không phải số liệu lấy từ Power BI thật.
- **Chưa chạy được trên Excel thật và chưa Refresh Power Query.** Môi trường dựng không có Excel. Query được nhúng bằng script (`build/pq_inject.py`).
  - Nếu Excel báo lỗi phần query: tạo lại từng query bằng Data > Get Data > From Other Sources > Blank Query > Advanced Editor, dán nội dung file trong `power_query/`. Đặt tên query trùng tên file.
  - Các query `KQ_*` tải vào bảng có sẵn cùng tên trên các sheet `DL_*` và `NGOAI_LE`. Các query còn lại chọn "Only Create Connection".
- Lần đầu Refresh: vào Data > Get Data > Query Options > Current Workbook > Privacy, chọn "Ignore the Privacy Levels". Query đọc thư mục sao kê theo đường dẫn trong bảng, nên Excel cần thiết lập này.

## Dựng lại file (khi sửa query hoặc báo cáo)

```
pip install openpyxl python-dateutil
python build/build_workbook.py <Du_lieu_dau_vao_du_no.xlsx> <bang_tra_no.md> <Bao_cao_du_no.xlsx> \
       --sao-ke HP:VPBank:<file sao kê> HP:NABank:<file sao kê>
```
