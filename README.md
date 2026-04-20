# AMZADS - Automated Advertising Trading System

Hệ thống quản lý và tối ưu hóa quảng cáo Amazon tự động dựa trên logic Trading. Dự án được thiết kế để xử lý dữ liệu hàng loạt (Bulk Operations) và điều phối chiến dịch theo thời gian thực.

## 📊 Cấu trúc Hệ thống (System Architecture)

Dự án được chia thành các module chuyên biệt theo quy trình vận hành:

* **A_create_campaign**: Khởi tạo cấu trúc chiến dịch dựa trên template chuẩn.
* **B_update_campaign**: Cập nhật thông số kỹ thuật cho các chiến dịch hiện hành.
* **C_mass_sop_factory**: Xử lý quy trình vận hành tiêu chuẩn (SOP) trên quy mô lớn.
* **D_auto_bulk_updater**: Engine tự động đọc và ghi dữ liệu Bulk File của Amazon.
* **F_core_engine**: Bộ xử lý trung tâm, điều phối logic giữa các module.
* **RULE&TEMPLATE**: Chứa các quy tắc tối ưu hóa (Bid, Placement, Budget).

## 🛠 Công nghệ sử dụng
* **Ngôn ngữ:** Python 3.x
* **Thư viện chính:** Pandas (Xử lý dữ liệu), Requests/Boto3 (Tương tác API).
* **Công cụ:** Amazon Advertising API, Bulk Sheets v3.

## 🚀 Hướng dẫn cài đặt

1. **Clone dự án:**
   ```bash
   git clone [https://github.com/maotieumao419-del/amzads.git](https://github.com/maotieumao419-del/amzads.git)
