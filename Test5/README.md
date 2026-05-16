# Test 5: Amazon Ads Automated Dashboard & Time-Series DB

## Bối cảnh
Test 5 là dự án nâng cấp và tự động hóa báo cáo phân tích hiệu suất Amazon Ads. Công cụ này xử lý dữ liệu từ nhiều file Bulk Operations, phân tích tự động dựa trên các quy tắc kinh doanh (Test 4 - Rule Engine) và xuất ra một Dashboard Excel trực quan với khả năng truy xuất lịch sử, đồ thị tương tác, và phân tích đa chiều.

Dự án này là Giai đoạn 1 (Phase 1: Excel Dashboard kết hợp SQLite) nhằm xây dựng kho dữ liệu và UI phục vụ nền tảng trước khi nâng cấp lên giao diện Web hoàn chỉnh ở Giai đoạn 2.

## Cấu trúc thư mục
- `input/`: Đặt các file `BulkSheetExport.xlsx` tải từ Amazon vào đây. Script sẽ tự động đọc, trích xuất metadata ngày tháng từ tên file và nạp vào DB.
- `output/`: Thư mục lưu kết quả file Dashboard hoàn thiện: `Amazon_Ads_Dashboard_{date}.xlsx`.
- `phase0/code/`: Chứa mã nguồn cho pipeline tạo báo cáo. Script chính điều phối là `main.py`.
- `phase_bo_sung/code/`: Chứa các script xử lý bổ sung như trích xuất tên file (metadata extraction).
- `database/`: Chứa `amazon_ads_history.db` (SQLite) và các file quản lý DB như `db_manager.py`. Hệ thống có chức năng chuẩn hoá, bù trừ (zero-fill) ngày thiếu để báo cáo chuỗi thời gian liên tục.
- `config/`: Chứa `Rule_Engine.json` và `Season_Calendar.json`.

## Các tính năng trên Dashboard
Script tự động sinh ra 1 file Excel duy nhất chứa các sheet hiển thị và các sheet ẩn chứa dữ liệu gốc (`RAW_TS`, `RAW_KW`):
1. **📊 OVERVIEW**: Thông số KPI toàn bộ tài khoản.
2. **🎯 CAMPAIGNS**: Thông tin chi tiết của tất cả các chiến dịch quảng cáo, gán Health Tag tự động, gắn biểu đồ Sparkline thu nhỏ của từng chiến dịch.
3. **🔑 KEYWORDS**: Hiệu suất từng Keyword/ASIN, và Đề xuất tăng/giảm Bid (theo Rule Engine).
4. **🔍 SEARCH TERMS**: Khai thác dữ liệu "Customer Search Term" thực tế.
5. **📅 DAILY TRENDS**: Báo cáo xu hướng chung của tài khoản (Account Level) theo từng ngày với đồ thị Line Chart.
6. **🔍 DEEP DIVE**: Trang phân tích sâu có tính năng tương tác:
   - Dùng Data Validation cho phép chọn linh hoạt **Campaign Name** hoặc **Campaign ID**.
   - Tự động lọc thông tin chi tiêu/doanh thu (Spend, Sales, ACOS) qua các khoảng thời gian kèm biểu đồ cột.
   - Sử dụng Dynamic Array (Hàm `FILTER`) để xuất **Danh sách lịch sử của các Keyword** thuộc Campaign đã chọn xuyên suốt các khoảng thời gian (Bulk files) khác nhau, đảm bảo báo cáo chi tiết đến mức từ khoá mà không làm đè lên biểu đồ (ngay cả khi Campaign có nhiều mốc thời gian).

## Hướng dẫn sử dụng
1. Đảm bảo đã tải và chép các file `.xlsx` vào thư mục `input/`.
2. Mở Command Prompt hoặc Terminal.
3. Chạy script chính để ingest data và render báo cáo:
   ```cmd
   f:\Minhpython\venv\Scripts\python.exe f:\Minhpython\Test5\phase0\code\main.py
   ```
4. Hệ thống sẽ:
   - Phân tích file bulk và đồng bộ vào SQLite database.
   - Bù dữ liệu rỗng và chuẩn bị Time-series (TS) data.
   - Render và định dạng các bảng tính/biểu đồ trên file Excel kết quả.
5. Mở thư mục `output/` để xem file báo cáo được sinh ra.
