# 📘 Hướng Dẫn Sử Dụng — Speech PowerPoint Controller

> Điều khiển PowerPoint bằng giọng nói tiếng Việt — hỗ trợ 2 chế độ: **Console** và **Dashboard Web**.

---

## 📋 Mục Lục

1. [Yêu cầu hệ thống](#-yêu-cầu-hệ-thống)
2. [Cài đặt](#-cài-đặt)
3. [Chuẩn bị file PowerPoint](#-chuẩn-bị-file-powerpoint)
4. [Chế độ 1 — Console (main.py)](#-chế-độ-1--console-mainpy)
5. [Chế độ 2 — Dashboard Web (dashboard_server.py)](#-chế-độ-2--dashboard-web-dashboard_serverpy)
6. [Bảng lệnh giọng nói](#-bảng-lệnh-giọng-nói)
7. [Quy trình demo chuẩn](#-quy-trình-demo-chuẩn)
8. [Xử lý lỗi thường gặp](#-xử-lý-lỗi-thường-gặp)
9. [Câu hỏi thường gặp (FAQ)](#-câu-hỏi-thường-gặp-faq)
10. [Cấu trúc Project](#-cấu-trúc-project)

---

## 💻 Yêu cầu hệ thống

| Yêu cầu | Chi tiết |
|----------|----------|
| **Hệ điều hành** | Windows 10/11 (bắt buộc — cần COM interface để điều khiển PowerPoint) |
| **Python** | 3.7 trở lên |
| **Microphone** | Đã kết nối và hoạt động |
| **Microsoft PowerPoint** | Đã cài đặt (Office 2016 trở lên khuyến nghị) |
| **Internet** | Cần có (Google Speech API yêu cầu kết nối mạng) |

---

## 🔧 Cài đặt

### Bước 1 — Tạo Virtual Environment

```bash
# Di chuyển vào thư mục project
cd SLP301_PowerpointController

# Tạo virtual environment
python -m venv venv

# Kích hoạt virtual environment
venv\Scripts\activate
```

### Bước 2 — Cài đặt thư viện

```bash
pip install -r requirements.txt
```

### Bước 3 — Xử lý nếu PyAudio bị lỗi

PyAudio thường gặp lỗi trên Windows vì thiếu thư viện portaudio. Thử các cách sau:

**Cách 1 — Dùng pipwin:**
```bash
pip install pipwin
pipwin install pyaudio
```

**Cách 2 — Download file .whl phù hợp:**
- Truy cập: https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio
- Tải file `.whl` phù hợp với phiên bản Python (vd: `PyAudio‑0.2.14‑cp311‑cp311‑win_amd64.whl`)
- Cài đặt:
```bash
pip install PyAudio-0.2.14-cp311-cp311-win_amd64.whl
```

### Bước 4 — Kiểm tra cài đặt

```bash
python -c "import speech_recognition; import win32com.client; print('Tat ca thu vien OK!')"
```

### Bước 5 — Test microphone

```bash
python test_microphone.py
```

Nói thử vài câu tiếng Việt. Nếu chương trình in ra đúng nội dung bạn nói → microphone đã sẵn sàng.

---

## 📂 Chuẩn bị file PowerPoint

1. Tạo hoặc copy bất kỳ file PowerPoint nào (`.pptx` hoặc `.ppt`)
2. Tạo thư mục `presentations/` trong thư mục gốc project (nếu chưa có)
3. Đặt file vào thư mục `presentations/`

```
SLP301_PowerpointController/
├── presentations/            ← Đặt file .pptx ở đây
│   ├── bai-thuyet-trinh.pptx
│   └── demo.pptx
├── main.py
├── dashboard_server.py
└── ...
```

> **Lưu ý:** Chương trình sẽ quét thư mục `presentations/` khi bạn nói **"tìm file"**.

---

## 🖥 Chế độ 1 — Console (`main.py`)

Chạy toàn bộ trên terminal, phù hợp để kiểm tra nhanh.

### Khởi chạy

```bash
python main.py
```

### Giao diện console

```
============================================================
🎤 SPEECH POWERPOINT CONTROLLER
============================================================

🎯 Sẵn sàng nhận lệnh giọng nói...
============================================================

🎤 Đang lắng nghe... (hãy nói lệnh)
🔄 Đang xử lý...
📝 Đã nhận: 'mở powerpoint'
✅ Lệnh: open_powerpoint
✔️  Đã mở PowerPoint
------------------------------------------------------------
```

### Dừng chương trình

- Nói **"tắt chương trình"** hoặc nhấn **Ctrl+C**.

---

## 🌐 Chế độ 2 — Dashboard Web (`dashboard_server.py`)

Giao diện web real-time, hiển thị đầy đủ trạng thái hệ thống — **tối ưu cho demo và quay video**.

### Khởi chạy

```bash
python dashboard_server.py
```

### Mở Dashboard

Sau khi server khởi động, mở trình duyệt và truy cập:

```
http://localhost:5000
```

### Giao diện Dashboard

Dashboard hiển thị real-time:

| Khu vực | Nội dung |
|---------|----------|
| **Microphone Status** | Trạng thái mic: Đang nghe / Đang xử lý / Chờ / Lỗi + animation sóng âm |
| **PowerPoint Status** | Trạng thái PPT: Chưa kết nối / Đã mở / Đang trình chiếu + tên file |
| **Current Command** | Lệnh vừa nhận dạng + action + kết quả (thành công/thất bại) |
| **Current Slide** | Slide hiện tại / tổng số slide + thanh tiến trình |
| **Command History** | Lịch sử toàn bộ lệnh đã thực thi (tối đa 50 entries) |
| **Log Panel** | Log real-time, phân màu theo mức độ (Info/Warning/Error), có bộ lọc |

### Dừng server

Nhấn **Ctrl+C** trên terminal đang chạy server.

---

## 🎤 Bảng lệnh giọng nói

### Khởi động

| Lệnh nói | Hành động | Ghi chú |
|-----------|-----------|---------|
| "mở powerpoint" | Mở ứng dụng Microsoft PowerPoint | Cần nói đầu tiên |
| "khởi động powerpoint" | (tương tự) | Từ khóa thay thế |

### Quản lý file

| Lệnh nói | Hành động | Ghi chú |
|-----------|-----------|---------|
| "tìm file" | Quét thư mục `presentations/` và hiển thị danh sách | Bắt buộc trước khi chọn file |
| "số 1", "số 2", ... | Chọn file theo số thứ tự trong danh sách | Nói sau "tìm file" |
| "số một", "số hai", ... | (tương tự, hỗ trợ số bằng chữ) | Hỗ trợ 1–9 |
| "mở file" | Mở file PowerPoint đã chọn | Nói sau khi đã chọn file |

### Điều khiển trình chiếu

| Lệnh nói | Hành động | Từ khóa thay thế |
|-----------|-----------|-------------------|
| "trình chiếu" | Bắt đầu chế độ Slide Show | "bắt đầu", "chạy" |
| "tiếp tục" | Chuyển sang slide tiếp theo | "next", "kế tiếp", "tiếp" |
| "lùi lại" | Quay lại slide trước | "back", "quay lại", "trước" |
| "slide 5" | Chuyển đến slide số 5 | "trang 5" |
| "slide ba" | Chuyển đến slide số 3 | Hỗ trợ số bằng chữ (1–9) |

### Kết thúc

| Lệnh nói | Hành động | Từ khóa thay thế |
|-----------|-----------|-------------------|
| "thoát" | Thoát chế độ trình chiếu (Slide Show) | "dừng", "kết thúc" |
| "tắt chương trình" | Đóng PowerPoint + thoát ứng dụng | "đóng chương trình" |

---

## 🎯 Quy trình demo chuẩn

Thực hiện tuần tự theo các bước sau:

### Bước 1 — Khởi động

```
Nói: "mở powerpoint"
→ PowerPoint sẽ mở lên
```

### Bước 2 — Tìm file

```
Nói: "tìm file"
→ Hiển thị danh sách file trong thư mục presentations/
   1. bai-thuyet-trinh.pptx
   2. demo.pptx
```

### Bước 3 — Chọn file

```
Nói: "số 1"
→ Đã chọn file: bai-thuyet-trinh.pptx
```

### Bước 4 — Mở file

```
Nói: "mở file"
→ File PowerPoint được mở
```

### Bước 5 — Bắt đầu trình chiếu

```
Nói: "trình chiếu"
→ Slide Show bắt đầu
```

### Bước 6 — Điều khiển slide

```
Nói: "tiếp tục"    → Chuyển sang slide tiếp theo
Nói: "slide 3"     → Nhảy đến slide 3
Nói: "lùi lại"     → Quay lại slide trước
```

### Bước 7 — Kết thúc

```
Nói: "thoát"              → Thoát chế độ trình chiếu
Nói: "tắt chương trình"   → Đóng PowerPoint + thoát ứng dụng
```

> **💡 Mẹo khi demo:**
> - Nói to, rõ, tự nhiên — tránh nói quá nhanh
> - Chờ chương trình hiển thị "Đang lắng nghe..." rồi mới nói
> - Nếu lệnh không nhận dạng được, đợi và nói lại
> - Đảm bảo internet ổn định (Google Speech API cần mạng)

---

## 🛠 Xử lý lỗi thường gặp

### 1. PyAudio không cài được

```
error: Microsoft Visual C++ 14.0 or greater is required
```

**Giải pháp:** Xem phần [Cài đặt → Bước 3](#bước-3--xử-lý-nếu-pyaudio-bị-lỗi)

---

### 2. Không nhận diện được giọng nói

**Nguyên nhân có thể:**
- Microphone chưa kết nối hoặc bị tắt
- Nói quá nhỏ hoặc có nhiều tiếng ồn
- Mất kết nối internet

**Giải pháp:**
- Kiểm tra microphone trong Windows Settings → Sound → Input
- Nói to, rõ ràng hơn
- Tăng timeout trong `config/settings.py`:

```python
SPEECH_CONFIG = {
    'language': 'vi-VN',
    'timeout': 10,           # Tăng lên 10 giây
    'phrase_time_limit': 8,  # Tăng lên 8 giây
}
```

---

### 3. PowerPoint không mở được

```
Lỗi: Không thể mở PowerPoint
```

**Giải pháp:**
- Đảm bảo Microsoft PowerPoint đã cài đặt
- Kiểm tra `pywin32`: `pip install pywin32`
- Thử mở PowerPoint thủ công trước

---

### 4. Không tìm thấy file PowerPoint

```
Không tìm thấy file PowerPoint trong thư mục presentations
```

**Giải pháp:**
- Tạo thư mục `presentations/` trong thư mục gốc project
- Đặt file `.pptx` hoặc `.ppt` vào thư mục đó
- Đảm bảo file không bị ẩn

---

### 5. Dashboard không hiển thị

```
ModuleNotFoundError: No module named 'flask'
```

**Giải pháp:**
```bash
pip install flask flask-socketio
```

---

## ❓ Câu hỏi thường gặp (FAQ)

**Q: Chương trình có chạy offline không?**
> A: Không. Google Speech API yêu cầu kết nối internet để nhận dạng giọng nói.

**Q: Hỗ trợ những ngôn ngữ nào?**
> A: Hiện tại chỉ hỗ trợ tiếng Việt (`vi-VN`). Có thể thay đổi trong `config/settings.py`.

**Q: Chạy trên macOS/Linux được không?**
> A: Không. PowerPoint COM interface (`win32com`) chỉ hoạt động trên Windows.

**Q: Độ trễ bao nhiêu?**
> A: Khoảng 2–3 giây cho mỗi lệnh (phụ thuộc tốc độ mạng).

**Q: Khác nhau giữa `main.py` và `dashboard_server.py`?**
> A: Cả hai đều chạy cùng logic (nghe → parse → thực thi). `main.py` hiển thị trên console, `dashboard_server.py` mở Dashboard web real-time tại `http://localhost:5000`.

**Q: Khác nhau giữa "số \<n\>" và "slide \<n\>"?**
> A: "số \<n\>" dùng để **chọn file** trong danh sách. "slide \<n\>" dùng để **chuyển slide** trong trình chiếu.

---

## 📁 Cấu trúc Project

```
SLP301_PowerpointController/
│
├── main.py                       # Entry point — chế độ Console
├── dashboard_server.py           # Entry point — chế độ Dashboard Web
├── requirements.txt              # Danh sách thư viện Python
├── test_microphone.py            # Script test microphone
│
├── speech/                       # Module nhận dạng giọng nói
│   ├── __init__.py
│   └── recognizer.py             #   SpeechRecognizer class
│
├── parser/                       # Module phân tích lệnh
│   ├── __init__.py
│   └── command_parser.py         #   CommandParser class
│
├── controller/                   # Module điều khiển PowerPoint
│   ├── __init__.py
│   └── powerpoint_controller.py  #   PowerPointController class
│
├── config/                       # Cấu hình toàn cục
│   ├── __init__.py
│   └── settings.py               #   SPEECH_CONFIG, COMMANDS, ...
│
├── utils/                        # Tiện ích chung
│   ├── __init__.py
│   └── logger.py                 #   Logging system
│
├── dashboard/                    # Giao diện Dashboard Web
│   ├── index.html                #   Trang Dashboard (HTML + JS)
│   └── style.css                 #   Dark theme stylesheet
│
├── presentations/                # Thư mục chứa file PowerPoint
│   └── (đặt file .pptx ở đây)
│
├── ARCHITECTURE.md               # Kiến trúc hệ thống
├── USER_GUIDE.md                 # ← Bạn đang đọc file này
└── README.md                     # Tổng quan project
```

---

> **SLP301m — Speech PowerPoint Controller Project**
