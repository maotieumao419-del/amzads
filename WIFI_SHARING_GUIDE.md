# 🌐 HƯỚNG DẪN CHIA SẺ DASHBOARD CHO NGƯỜI CÙNG MẠNG WIFI

Tài liệu này hướng dẫn bạn cách khởi chạy hệ thống Ads Trading System để bất kỳ ai đang dùng **chung một mạng WiFi** (như đồng nghiệp trong văn phòng, sếp, hoặc chính bạn dùng điện thoại) đều có thể truy cập vào xem và bấm nút chạy code.

> [!NOTE]
> Tôi đã cập nhật một chút code ở Frontend để đảm bảo hệ thống có thể hoạt động trơn tru qua môi trường mạng LAN. Bạn chỉ cần làm đúng theo các bước dưới đây.

---

## Bước 1: Tìm địa chỉ IP máy tính của bạn
Đây là "địa chỉ nhà" của máy tính bạn trong mạng WiFi. Người khác sẽ nhập địa chỉ này vào trình duyệt để vào Web.

1. Mở một **Terminal mới** (Powershell hoặc Command Prompt).
2. Gõ lệnh sau và nhấn Enter:
   ```powershell
   ipconfig
   ```
3. Tìm đến dòng có chữ `IPv4 Address`. Nó sẽ có dạng các con số như: `192.168.1.45` hoặc `10.0.0.x`.
   > **Lưu ý:** Hãy ghi nhớ hoặc copy dãy số này lại. (Ví dụ trong bài hướng dẫn này tôi sẽ dùng `192.168.1.45`).

---

## Bước 2: Bật Backend (Máy chủ xử lý Python)
Backend làm nhiệm vụ xử lý data và chạy các luồng tự động hóa.
1. Mở một Terminal trong IDE (nhớ kích hoạt `venv`).
2. Khởi chạy bằng lệnh:
   ```powershell
   python UI\serve_webapp.py
   ```
3. *Đảm bảo Terminal báo: `Uvicorn running on http://0.0.0.0:8000` (Nghĩa là nó đã sẵn sàng đón khách từ xa).*

---

## Bước 3: Bật Frontend (Giao diện Web) ở chế độ Mạng LAN
Đây là bước cực kỳ quan trọng để cho phép người khác vào được trang Web.

1. Mở một **Terminal thứ 2**.
2. Di chuyển vào thư mục web:
   ```powershell
   cd UI\webapp
   ```
3. Thay vì chạy lệnh bình thường, bạn hãy chạy lệnh sau để mở khóa mạng:
   ```powershell
   npm run dev -- --host
   ```
4. *Lúc này, trên màn hình Terminal sẽ hiện ra 2 dòng Network. Một dòng là Local (chỉ mình bạn), một dòng là Network (dành cho người khác).*
   ```text
   ➜  Local:   http://localhost:5173/
   ➜  Network: http://192.168.1.45:5173/  <-- Đây là link để chia sẻ!
   ```

---

## Bước 4: Chia sẻ và Tận hưởng! 🎉
- Bạn chỉ cần copy dòng link **Network** ở trên (`http://192.168.1.45:5173/`) và gửi qua Zalo/Tin nhắn cho người đồng nghiệp.
- Người đó (dùng laptop hoặc điện thoại có kết nối chung WiFi) nhấp vào link là sẽ lập tức mở ra trang Dashboard y hệt máy bạn.
- Khi người đó bấm nút **"🚀 Run Full Code"**, lệnh sẽ được gửi thẳng qua WiFi về máy tính của bạn, và máy bạn sẽ tự động chạy file Python. Đỉnh chưa!

> [!WARNING]
> **Vấn đề thường gặp:** Nếu người khác báo không vào được dù đã chung WiFi, 99% nguyên nhân là do **Tường lửa (Firewall)** trên Windows của bạn đang chặn kết nối. Hãy tắt tạm tường lửa hoặc cho phép (allow) cổng `5173` và `8000` trong phần cài đặt Windows Defender Firewall nhé.
