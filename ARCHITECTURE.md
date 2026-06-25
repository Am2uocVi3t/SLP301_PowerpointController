# 🏗️ Kiến trúc hệ thống - Speech PowerPoint Controller

## 📊 Sơ đồ tổng quan

```
┌─────────────────────────────────────────────────────────────────┐
│                         USER (Người dùng)                        │
│                     🗣️ Nói lệnh bằng giọng nói                  │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                    🎤 MICROPHONE (Hardware)                      │
│                   Capture audio từ người dùng                    │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│              📦 SPEECH MODULE (speech/recognizer.py)             │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  1. Ghi âm từ microphone (PyAudio)                       │  │
│  │  2. Gửi audio lên Google Speech API                      │  │
│  │  3. Nhận text tiếng Việt                                 │  │
│  └──────────────────────────────────────────────────────────┘  │
│              Output: text = "mở powerpoint"                      │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│            🔍 PARSER MODULE (parser/command_parser.py)           │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  1. Nhận text từ Speech Module                           │  │
│  │  2. So khớp với danh sách keywords                       │  │
│  │  3. Tạo command object có cấu trúc                       │  │
│  └──────────────────────────────────────────────────────────┘  │
│       Output: {action: "open_powerpoint", params: {}}            │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│       🎮 CONTROLLER MODULE (controller/powerpoint_controller.py) │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  1. Nhận command object                                  │  │
│  │  2. Gọi method tương ứng                                 │  │
│  │  3. Tương tác với PowerPoint qua COM interface           │  │
│  └──────────────────────────────────────────────────────────┘  │
│              Output: {success: True, message: "..."}             │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                  💻 MICROSOFT POWERPOINT (COM)                   │
│                    Thực hiện thao tác thực tế                    │
└─────────────────────────────────────────────────────────────────┘
```

## 🔄 Luồng xử lý chi tiết

### 1️⃣ Main Loop (main.py)

```python
while True:
    # STEP 1: Lắng nghe
    text = speech_recognizer.listen()
    
    # STEP 2: Phân tích
    command = command_parser.parse(text)
    
    # STEP 3: Thực thi
    result = ppt_controller.execute(command)
    
    # STEP 4: Hiển thị kết quả
    print(result['message'])
```

### 2️⃣ Speech Recognition Flow

```
Microphone Input
       ↓
   PyAudio capture audio
       ↓
   Convert to AudioData object
       ↓
   Send to Google Speech API
       ↓
   Receive text response
       ↓
   Return lowercase text
```

### 3️⃣ Command Parsing Flow

```
Input: "mở powerpoint"
       ↓
   Lowercase & trim: "mở powerpoint"
       ↓
   Loop through COMMANDS config
       ↓
   Check if any keyword matches
       ↓
   Found: keywords=['mở powerpoint', 'mo powerpoint']
       ↓
   Return: {action: 'open_powerpoint', params: {}}
```

### 4️⃣ Controller Execution Flow

```
Input: {action: 'open_powerpoint', params: {}}
       ↓
   Map action to method: _open_powerpoint()
       ↓
   Execute method:
       ↓
   win32com.client.Dispatch("PowerPoint.Application")
       ↓
   powerpoint.Visible = True
       ↓
   Return: {success: True, message: 'Đã mở PowerPoint'}
```

## 📂 Cấu trúc module chi tiết

### SPEECH MODULE
```
speech/
├── __init__.py
└── recognizer.py
    └── class SpeechRecognizer
        ├── __init__()           # Khởi tạo recognizer và mic
        └── listen()             # Lắng nghe và nhận dạng
```

**Nhiệm vụ**: Input audio → Output text

### PARSER MODULE
```
parser/
├── __init__.py
└── command_parser.py
    └── class CommandParser
        ├── __init__()                    # Load COMMANDS config
        ├── parse(text)                   # Parse text thành command
        └── _extract_slide_number(text)   # Extract số slide
```

**Nhiệm vụ**: Input text → Output structured command

### CONTROLLER MODULE
```
controller/
├── __init__.py
└── powerpoint_controller.py
    └── class PowerPointController
        ├── __init__()              # Khởi tạo
        ├── execute(command)        # Dispatcher chính
        ├── _open_powerpoint()      # Mở PowerPoint
        ├── _find_file()            # Tìm và mở file
        ├── _goto_slide()           # Chuyển slide
        ├── _start_slideshow()      # Bắt đầu trình chiếu
        ├── _next_slide()           # Slide tiếp
        ├── _previous_slide()       # Slide trước
        └── _exit()                 # Thoát
```

**Nhiệm vụ**: Input command → Output execution result

### CONFIG MODULE
```
config/
├── __init__.py
└── settings.py
    ├── SPEECH_CONFIG           # Cấu hình speech recognition
    ├── POWERPOINT_CONFIG       # Cấu hình PowerPoint
    ├── COMMANDS                # Định nghĩa các lệnh
    └── LOGGING_CONFIG          # Cấu hình logging
```

**Nhiệm vụ**: Tập trung tất cả cấu hình

### UTILS MODULE
```
utils/
├── __init__.py
└── logger.py
    ├── setup_logger()          # Thiết lập logger
    └── get_logger()            # Lấy logger instance
```

**Nhiệm vụ**: Logging, utilities chung

## 🔗 Dependency Graph

```
main.py
  ├─→ speech.recognizer
  │     ├─→ speech_recognition
  │     └─→ config.settings
  │
  ├─→ parser.command_parser
  │     └─→ config.settings
  │
  └─→ controller.powerpoint_controller
        ├─→ win32com.client
        └─→ config.settings
```

## 💾 Data Flow Example

### Ví dụ: Lệnh "mở powerpoint"

```
USER speaks: "mở powerpoint"
          ↓
[MICROPHONE] captures audio wave
          ↓
[SpeechRecognizer.listen()]
  • PyAudio records audio
  • Sends to Google API
  • Returns: "mở powerpoint"
          ↓
[CommandParser.parse("mở powerpoint")]
  • Checks keywords: ['mở powerpoint', 'mo powerpoint']
  • Matches! action = 'open_powerpoint'
  • Returns: {action: 'open_powerpoint', params: {}}
          ↓
[PowerPointController.execute(command)]
  • Maps to _open_powerpoint()
  • Calls win32com.client.Dispatch("PowerPoint.Application")
  • Sets Visible = True
  • Returns: {success: True, message: 'Đã mở PowerPoint'}
          ↓
[MAIN LOOP] prints: "✔️ Đã mở PowerPoint"
```

## 🎯 Design Patterns

### 1. Command Pattern
- Text → Command Object → Execute
- Dễ dàng thêm lệnh mới

### 2. Strategy Pattern
- Mỗi action là một strategy (_open_powerpoint, _find_file, etc.)
- Runtime dispatch dựa trên action

### 3. Singleton Pattern
- Logger singleton để tránh tạo nhiều instance

### 4. Facade Pattern
- PowerPointController che giấu phức tạp của COM interface

## 🔐 Error Handling Strategy

```
Try-Except at 3 levels:

1. SPEECH LEVEL
   • WaitTimeoutError → Không nghe thấy
   • UnknownValueError → Không nhận dạng được
   • RequestError → Lỗi API

2. PARSER LEVEL
   • No match → Return None

3. CONTROLLER LEVEL
   • COM Error → Catch và return error message
   • File not found → Return error message
```

## 📈 Extensibility

### Thêm lệnh mới (3 bước):

1. **Thêm vào config/settings.py**
```python
{
    'action': 'close_presentation',
    'keywords': ['đóng file', 'dong file'],
    'description': 'Đóng presentation'
}
```

2. **Thêm method vào controller**
```python
def _close_presentation(self, params):
    self.presentation.Close()
    return {'success': True, 'message': 'Đã đóng file'}
```

3. **Update action_map**
```python
action_map = {
    ...
    'close_presentation': self._close_presentation
}
```

---

**Kiến trúc này đảm bảo:**
- ✅ Separation of Concerns
- ✅ Easy to Test
- ✅ Easy to Extend
- ✅ Maintainable
