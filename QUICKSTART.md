# ⚡ QUICKSTART - Chạy ngay trong 5 phút

## Bước 1: Cài đặt (2 phút)

```bash
# Clone hoặc giải nén project
cd speech-powerpoint-controller

# Tạo virtual environment
python -m venv venv

# Activate
venv\Scripts\activate

# Cài thư viện
pip install -r requirements.txt
```

**Nếu PyAudio lỗi:**
```bash
pip install pipwin
pipwin install pyaudio
```

## Bước 2: Chuẩn bị file PowerPoint (1 phút)

1. Tạo hoặc copy file PowerPoint bất kỳ (`.pptx`)
2. Đặt vào thư mục `demo/`
3. Ví dụ: `demo/my-presentation.pptx`

## Bước 3: Test microphone (1 phút)

```bash
python test_microphone.py
```

Nói thử vài câu tiếng Việt. Nếu nhận dạng được → OK!

## Bước 4: Chạy chương trình (1 phút)

```bash
python main.py
```

## Thử các lệnh:

```
1. Nói: "mở powerpoint"      → PowerPoint sẽ mở
2. Nói: "tìm file"           → File .pptx trong demo/ sẽ mở
3. Nói: "trình chiếu"        → Bắt đầu slideshow
4. Nói: "tiếp tục"           → Chuyển slide
5. Nói: "slide hai"          → Nhảy đến slide 2
6. Nói: "lùi lại"            → Quay lại slide trước
7. Nói: "thoát"              → Thoát
```

## Xong! 🎉

Nếu gặp lỗi, xem file `TUTORIAL.md` để biết cách khắc phục chi tiết.

## Các lệnh hỗ trợ

| Lệnh | Từ khóa thay thế |
|------|------------------|
| Mở PowerPoint | "khởi động powerpoint" |
| Tìm file | "mở file", "tìm kiếm file" |
| Slide 1-9 | "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" |
| Trình chiếu | "bắt đầu", "chạy" |
| Tiếp tục | "next", "kế tiếp" |
| Lùi lại | "back", "quay lại", "trước" |
| Thoát | "dừng", "kết thúc", "exit" |
