# 📖 Hướng dẫn chi tiết - Speech PowerPoint Controller

## 🎓 Dành cho sinh viên demo đồ án

### 1️⃣ CHUẨN BỊ TRƯỚC KHI DEMO

#### Bước 1: Cài đặt môi trường

```bash
# 1. Tạo virtual environment
python -m venv venv

# 2. Activate
venv\Scripts\activate

# 3. Cài thư viện
pip install -r requirements.txt

# 4. Kiểm tra cài đặt
python -c "import speech_recognition; import win32com.client; print('OK')"
```

#### Bước 2: Chuẩn bị file PowerPoint

1. Tạo file PowerPoint đơn giản (5-10 slides)
2. Đặt tên: `demo.pptx`
3. Copy vào thư mục `demo/`

#### Bước 3: Test microphone

```python
# test_mic.py
import speech_recognition as sr

recognizer = sr.Recognizer()
mic = sr.Microphone()

with mic as source:
    print("Nói thử...")
    audio = recognizer.listen(source, timeout=5)
    text = recognizer.recognize_google(audio, language='vi-VN')
    print(f"Nhận được: {text}")
```

### 2️⃣ QUY TRÌNH DEMO CHUẨN (5 phút)

#### Phần 1: Giới thiệu (1 phút)

"Em xin giới thiệu đồ án: Speech PowerPoint Controller
- Mục đích: Điều khiển PowerPoint bằng giọng nói
- Công nghệ: Python + SpeechRecognition + pywin32
- Tính năng: 7 lệnh giọng nói cơ bản"

#### Phần 2: Giải thích kiến trúc (1.5 phút)

```
1. SPEECH MODULE
   - Nhận audio từ microphone
   - Gửi lên Google Speech API
   - Trả về text tiếng Việt

2. PARSER MODULE
   - So khớp keywords trong text
   - Tạo command object có cấu trúc

3. CONTROLLER MODULE
   - Nhận command
   - Gọi COM interface của PowerPoint
   - Thực thi thao tác
```

#### Phần 3: Live Demo (2 phút)

```bash
# Chạy
python main.py

# Demo tuần tự:
1. "mở powerpoint"     → PowerPoint mở lên
2. "tìm file"          → File demo.pptx được mở
3. "trình chiếu"       → Slideshow bắt đầu
4. "tiếp tục"          → Chuyển slide 2
5. "slide ba"          → Nhảy đến slide 3
6. "lùi lại"           → Về slide 2
7. "thoát"             → Thoát chương trình
```

#### Phần 4: Code Walkthrough (0.5 phút)

Mở file quan trọng nhất và giải thích:

```python
# main.py - Vòng lặp chính
while True:
    text = speech_recognizer.listen()     # 1. Nghe
    command = command_parser.parse(text)  # 2. Parse
    result = ppt_controller.execute(cmd)  # 3. Execute
```

### 3️⃣ CÂU HỎI THƯỜNG GẶP KHI DEMO

**Q: Tại sao dùng Google API mà không dùng offline?**
A: Google API có độ chính xác cao cho tiếng Việt. Để offline có thể dùng Vosk hoặc Whisper (đã đề xuất trong phần cải tiến).

**Q: Làm sao xử lý nhiễu?**
A: Có dùng `adjust_for_ambient_noise()` trong recognizer.py để lọc nhiễu nền.

**Q: Nếu nói sai, có sửa được không?**
A: Hiện tại chưa có undo. Có thể cải tiến bằng cách lưu history và thêm lệnh "hoàn tác".

**Q: Chạy trên Mac/Linux được không?**
A: Không, vì PowerPoint COM chỉ có trên Windows. Nhưng có thể chuyển sang LibreOffice Impress (cross-platform).

**Q: Độ trễ bao nhiêu?**
A: Khoảng 2-3 giây (do gọi API). Có thể giảm bằng offline recognition.

### 4️⃣ XỬ LÝ LỖI PHỔ BIẾN

#### Lỗi 1: PyAudio không cài được

```bash
# Giải pháp
pip install pipwin
pipwin install pyaudio

# Hoặc
# Download từ: https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio
pip install PyAudio-0.2.11-cp310-cp310-win_amd64.whl
```

#### Lỗi 2: "No module named win32com"

```bash
pip install pywin32
```

#### Lỗi 3: Không nghe thấy giọng nói

```python
# Tăng timeout trong config/settings.py
SPEECH_CONFIG = {
    'timeout': 10,  # Tăng lên 10 giây
    'phrase_time_limit': 8
}
```

#### Lỗi 4: PowerPoint không mở được

```python
# Kiểm tra PowerPoint đã cài chưa
import win32com.client
try:
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    print("PowerPoint OK")
except:
    print("PowerPoint chưa cài hoặc lỗi")
```

### 5️⃣ ĐIỂM CỘNG KHI DEMO

✅ **Code sạch, có comment đầy đủ**
✅ **Kiến trúc modular, dễ hiểu**
✅ **Demo mượt, không lỗi**
✅ **Giải thích rõ ràng, tự tin**
✅ **Có logging, error handling**
✅ **README.md chi tiết**
✅ **Đề xuất cải tiến hợp lý**

### 6️⃣ CHECKLIST TRƯỚC KHI NỘP

- [ ] Code chạy được, không lỗi
- [ ] requirements.txt đầy đủ
- [ ] README.md đầy đủ thông tin
- [ ] Comment code đầy đủ
- [ ] Có file .gitignore
- [ ] Demo file PowerPoint sẵn sàng
- [ ] Test microphone trước
- [ ] Test tất cả lệnh
- [ ] Chuẩn bị trả lời câu hỏi
- [ ] Backup code (USB, GitHub)

### 7️⃣ TEMPLATE THUYẾT TRÌNH

```
1. GIỚI THIỆU (30s)
   - Tên đề tài
   - Mục đích
   - Công nghệ sử dụng

2. KIẾN TRÚC (1 phút)
   - Vẽ diagram: Mic → Speech → Parser → Controller → PowerPoint
   - Giải thích từng module

3. LIVE DEMO (2 phút)
   - Chạy và demo 7 lệnh
   - Comment trong khi demo

4. CODE REVIEW (1 phút)
   - Mở main.py
   - Giải thích workflow chính

5. CẢI TIẾN (30s)
   - Liệt kê 3-4 cải tiến có thể làm
   - Giải thích ngắn gọn

6. KẾT LUẬN (30s)
   - Tổng kết
   - Khó khăn gặp phải
   - Bài học rút ra
```

### 8️⃣ TIPS ĐỂ ĐIỂM CAO

1. **Tự tin**: Demo thành thạo, không nhìn cheat sheet
2. **Giải thích rõ**: Dùng từ ngữ chuyên môn đúng
3. **Xử lý lỗi**: Nếu demo lỗi, giải thích được lý do
4. **Trả lời tốt**: Hiểu sâu về code, không học vẹt
5. **Cải tiến**: Đề xuất cải tiến thực tế, khả thi

---

**Chúc bạn demo thành công! 🎉**
