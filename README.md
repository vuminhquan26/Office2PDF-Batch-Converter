# 📄 Office2PDF Batch Converter

Một công cụ mã nguồn mở nhỏ gọn, giao diện trực quan giúp chuyển đổi hàng loạt tài liệu Microsoft Word (`.doc`, `.docx`) và PowerPoint (`.ppt`, `.pptx`) sang định dạng PDF. 

Công cụ này chạy **100% Offline**, tương tác trực tiếp với lõi (COM object) của Microsoft Office đang cài trên máy để đảm bảo file PDF xuất ra giữ nguyên 100% định dạng gốc (không vỡ font, không lệch ảnh, không lỗi bảng biểu).

---

## ✨ Tính năng nổi bật

- **Chuyển đổi hàng loạt (Batch Convert):** Chọn hàng chục, hàng trăm file cùng lúc để tự động chuyển đổi.
- **Hỗ trợ đa định dạng:** Nhận diện và xử lý mượt mà cả file Word cũ/mới (`.doc`, `.docx`) và PowerPoint (`.ppt`, `.pptx`).
- **Giao diện siêu nhẹ:** Xây dựng bằng Tkinter, chiếm cực ít RAM và khởi động tức thì.
- **Hoạt động ngầm thông minh:** Ứng dụng thuật toán *Lazy Loading*, tận dụng tiến trình Word/PowerPoint chạy nền để tối ưu tốc độ, tự động dọn dẹp RAM sau khi hoàn tất.
- **Bảo mật tuyệt đối:** Không tải file lên web, mọi quá trình xử lý diễn ra cục bộ trên máy tính của bạn.

---

## 💻 Yêu cầu hệ thống

- Hệ điều hành: **Windows** (7 / 10 / 11).
- Bắt buộc phải cài đặt sẵn **Microsoft Office** (Word và PowerPoint) trên máy tính.
- Môi trường chạy code (Nếu chạy từ Source): **Python 3.x**.

---

## 🚀 Hướng dẫn sử dụng (Dành cho người dùng cuối)

Nếu bạn đã có file `DocToPDF.exe`, bạn không cần cài đặt gì thêm, chỉ cần mở lên và sử dụng:

1. **Bước 1: Thiết lập thư mục lưu PDF**
   - Click vào nút **"1. Thiết lập thư mục lưu PDF..."**
   - Chọn một thư mục trên máy tính mà bạn muốn chứa các file PDF sau khi chuyển đổi xong.
   
2. **Bước 2: Chọn file và Bắt đầu**
   - Click vào nút **"2. Chọn File Cần Chuyển & Bắt Đầu"**
   - Quét chọn một hoặc nhiều file `.doc`, `.docx`, `.ppt`, `.pptx` cần chuyển.
   - Ngồi nhâm nhi cà phê và xem thanh tiến trình (Progress bar) tự động chạy.
   
3. ⚠️ **Lưu ý quan trọng:** 
   - Ứng dụng sẽ xóa bộ nhớ điểm lưu khi đóng lại. Mỗi lần mở app, bạn cần thiết lập lại điểm lưu ở Bước 1.
   - Nếu app báo lỗi trong quá trình chạy, hãy mở Task Manager (`Ctrl + Shift + Esc`), tìm và End Task các tiến trình có tên `WINWORD.EXE` hoặc `POWERPNT.EXE` đang bị kẹt ẩn, sau đó thử lại.

---

## 🛠 Hướng dẫn cài đặt & Build (Dành cho Developer)

Nếu bạn muốn chỉnh sửa mã nguồn và tự đóng gói lại dự án:

### 1. Cài đặt thư viện
Clone repository này về máy, mở Terminal tại thư mục chứa source code và chạy:
```bash
pip install pywin32 pyinstaller
