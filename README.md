# 🎤 Speech PowerPoint Controller

Ứng dụng điều khiển PowerPoint bằng giọng nói - Đồ án môn học Python

## 📋 Giới thiệu

Speech PowerPoint Controller là một ứng dụng Python cho phép người dùng điều khiển PowerPoint bằng giọng nói tiếng Việt. Ứng dụng sử dụng microphone để nhận lệnh giọng nói, chuyển đổi thành text, và thực thi các thao tác tương ứng trên PowerPoint.

## ✨ Tính năng

- 🎙️ Nhận dạng giọng nói tiếng Việt
- 📂 Mở PowerPoint và tìm file tự động
- 🔢 Chuyển slide theo số (1-9)
- ▶️ Bắt đầu trình chiếu
- ⏭️ Chuyển slide tiếp theo
- ⏮️ Lùi lại slide trước
- 🚪 Thoát chương trình

## 🏗️ Cấu trúc Project


```
speech-powerpoint-controller/
├── main.py                          # Entry point chính
├── requirements.txt                 # Danh sách thư viện
├── README.md                        # Tài liệu hướng dẫn
│
├── speech/                          # Module xử lý giọng nói
│   ├── __init__.py
│   └── recognizer.py               # Nhận dạng giọng nói
│
├── parser/                          # Module phân tích lệnh
│   ├── __init__.py
│   └── command_parser.py           # Parse text thành lệnh
│
├── controller/                      # Module điều khiển PowerPoint
│   ├── __init__.py
│   └── powerpoint_controller.py    # Điều khiển PPT qua COM
│
├── config/                          # Module cấu hình
│   ├── __init__.py
│   └── settings.py                 # Các thiết lập
│
└── utils/                           # Module tiện ích
    ├── __init__.py
    └── logger.py                    # Logging system
```

## 🚀 Cài đặt

### 1. Yêu cầu hệ thống

- **Hệ điều hành**: Windows (vì cần tương tác với PowerPoint)
- **Python**: 3.7 trở lên
- **Microphone**: Microphone hoạt động tốt
- **Microsoft PowerPoint**: Đã cài đặt trên máy

### 2. Cài đặt Python packages


```bash
# Tạo virtual environment (khuyến nghị)
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate

# Cài đặt các thư viện
pip install -r requirements.txt
```

### 3. Lưu ý cài đặt PyAudio

PyAudio có thể gặp lỗi khi cài đặt. Nếu gặp lỗi, thử các cách sau:

**Cách 1**: Cài đặt từ wheel file
```bash
pip install pipwin
pipwin install pyaudio
```

**Cách 2**: Download wheel file phù hợp với Python version
- Truy cập: https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio
- Download file .whl phù hợp
- Cài đặt: `pip install PyAudio-xxx.whl`

## 📖 Hướng dẫn sử dụng

### 1. Chuẩn bị

- Đảm bảo microphone đã kết nối và hoạt động
- Chuẩn bị file PowerPoint (.pptx) trong thư mục demo hoặc presentations

### 2. Chạy chương trình


```bash
python main.py
```

### 3. Các lệnh giọng nói

| Lệnh | Mô tả | Ví dụ |
|------|-------|-------|
| **"mở powerpoint"** | Mở ứng dụng PowerPoint | "mở powerpoint" |
| **"tìm file"** | Tìm và mở file PowerPoint | "tìm file" |
| **"một/hai/ba"** | Chuyển đến slide 1/2/3 | "slide ba" |
| **"trình chiếu"** | Bắt đầu trình chiếu | "bắt đầu trình chiếu" |
| **"tiếp tục"** | Slide tiếp theo | "tiếp tục", "next" |
| **"lùi lại"** | Slide trước đó | "lùi lại", "back" |
| **"thoát"** | Thoát chương trình | "thoát" |

### 4. Quy trình demo chuẩn

1. Chạy chương trình: `python main.py`
2. Nói: **"mở powerpoint"** - Mở PowerPoint
3. Nói: **"tìm file"** - Mở file presentation
4. Nói: **"trình chiếu"** - Bắt đầu slideshow
5. Nói: **"tiếp tục"** - Chuyển slide tiếp theo
6. Nói: **"slide hai"** - Chuyển đến slide 2
7. Nói: **"lùi lại"** - Quay lại slide trước
8. Nói: **"thoát"** - Đóng chương trình

## 📝 Giải thích các file chính



### `main.py`
- **Chức năng**: Entry point của chương trình
- **Nhiệm vụ**: Khởi tạo các module và chạy vòng lặp chính
- **Workflow**: Listen → Parse → Execute → Repeat

### `speech/recognizer.py`
- **Chức năng**: Nhận dạng giọng nói từ microphone
- **Công nghệ**: SpeechRecognition + Google Speech API
- **Input**: Audio từ microphone
- **Output**: Text tiếng Việt

### `parser/command_parser.py`
- **Chức năng**: Phân tích text thành lệnh có cấu trúc
- **Input**: Text từ recognizer
- **Output**: Dictionary {action, params}
- **Logic**: So khớp keywords với danh sách lệnh

### `controller/powerpoint_controller.py`
- **Chức năng**: Điều khiển PowerPoint thông qua COM interface
- **Công nghệ**: pywin32 (win32com.client)
- **Các thao tác**: Open, Find, Navigate, Slideshow, Exit

### `config/settings.py`
- **Chức năng**: Tập trung tất cả cấu hình
- **Nội dung**: Cấu hình speech, PowerPoint, commands, logging
- **Lợi ích**: Dễ dàng điều chỉnh tham số

### `utils/logger.py`
- **Chức năng**: Ghi log hoạt động và lỗi
- **Output**: File app.log
- **Lợi ích**: Debug và theo dõi hoạt động

## 🔧 Troubleshooting



### Lỗi: PyAudio không cài được
- Dùng pipwin hoặc download wheel file (xem phần cài đặt)

### Lỗi: Không nhận diện giọng nói
- Kiểm tra microphone có hoạt động không
- Thử nói to và rõ hơn
- Kiểm tra kết nối internet (Google API cần internet)

### Lỗi: PowerPoint không mở
- Đảm bảo đã cài PowerPoint trên máy Windows
- Kiểm tra pywin32 đã cài đúng chưa

### Lỗi: Không tìm thấy file
- Đặt file .pptx vào thư mục `demo/` hoặc `presentations/`
- Hoặc đặt ngay trong thư mục gốc project

## 💡 Đề xuất cải tiến nâng cao

### Cải tiến cơ bản (7-8 điểm)

1. **Offline Speech Recognition**
   - Thêm Vosk hoặc Whisper để nhận dạng offline
   - Không cần internet, tăng tốc độ

2. **Thêm lệnh nâng cao**
   - "slide cuối cùng", "slide đầu tiên"
   - "tạm dừng", "tiếp tục" trong slideshow
   - "đóng PowerPoint"

3. **GUI đơn giản**
   - Dùng Tkinter tạo giao diện hiển thị trạng thái
   - Hiển thị slide hiện tại, lệnh vừa thực hiện



### Cải tiến nâng cao (8-9 điểm)

4. **Custom Wake Word**
   - Thêm wake word "Hey PowerPoint" trước lệnh
   - Dùng Porcupine hoặc Snowboy

5. **Multi-language Support**
   - Hỗ trợ cả tiếng Anh và tiếng Việt
   - Tự động detect ngôn ngữ

6. **Voice Feedback**
   - Text-to-Speech phản hồi lại người dùng
   - Dùng pyttsx3 hoặc gTTS

7. **Gesture Control Integration**
   - Kết hợp với camera để nhận dạng cử chỉ tay
   - Dùng MediaPipe hoặc OpenCV

### Cải tiến xuất sắc (9-10 điểm)

8. **Machine Learning Intent Classification**
   - Train model ML để phân loại intent chính xác hơn
   - Dùng sklearn hoặc TensorFlow

9. **Web Interface**
   - Tạo web app với Flask/FastAPI
   - Điều khiển từ xa qua browser

10. **Mobile App Integration**
    - Tạo app mobile (React Native/Flutter)
    - Điều khiển PowerPoint từ điện thoại

11. **Cloud Deployment**
    - Deploy lên cloud (AWS/Azure)
    - Tích hợp với Microsoft Teams/Zoom



## 📚 Thư viện sử dụng

| Thư viện | Phiên bản | Mục đích |
|----------|-----------|----------|
| **SpeechRecognition** | 3.10.1 | Nhận dạng giọng nói |
| **PyAudio** | 0.2.14 | Xử lý microphone input |
| **pywin32** | 306 | Điều khiển PowerPoint qua COM |

## 🎯 Điểm mạnh của thiết kế

1. **Modular Architecture**: Code tách bạch rõ ràng theo chức năng
2. **Easy to Extend**: Dễ dàng thêm lệnh mới trong `config/settings.py`
3. **Error Handling**: Xử lý lỗi đầy đủ, có logging
4. **Vietnamese Support**: Hỗ trợ tiếng Việt tốt với nhiều biến thể
5. **Production-Ready**: Có logging, config, error handling

## 👨‍💻 Tác giả

Đồ án môn học SLP301 - Speech PowerPoint Controller
Huỳnh Quốc Việt

## 📄 License

MIT License - Free to use for educational purposes

---

**Lưu ý**: Đây là project demo cho mục đích học tập. Để sử dụng trong thực tế, cần cải tiến thêm về bảo mật, performance và UX.
