"""
Settings - Cấu hình toàn cục cho ứng dụng
Chứa tất cả các tham số có thể điều chỉnh
"""

# Cấu hình Speech Recognition
SPEECH_CONFIG = {
    'language': 'vi-VN',  # Ngôn ngữ tiếng Việt
    'timeout': 5,  # Thời gian chờ nghe (giây)
    'phrase_time_limit': 5,  # Thời gian tối đa cho một câu (giây)
}

# Cấu hình PowerPoint
POWERPOINT_CONFIG = {
    'default_path': '.',  # Thư mục mặc định tìm file
}

# Định nghĩa các lệnh và từ khóa
COMMANDS = [
    {
        'action': 'open_powerpoint',
        'keywords': ['mở powerpoint', 'mo powerpoint', 'khởi động powerpoint', 'khoi dong powerpoint'],
        'description': 'Mở ứng dụng PowerPoint'
    },
    {
        'action': 'find_file',
        'keywords': ['tìm file', 'tim file', 'mở file', 'mo file', 'tìm kiếm file'],
        'description': 'Tìm và mở file PowerPoint'
    },
    {
        'action': 'goto_slide',
        'keywords': ['một', 'mot', 'hai', 'ba', 'bốn', 'bon', 'tư', 'tu', 'năm', 'nam', 
                    'sáu', 'sau', 'bảy', 'bay', 'tám', 'tam', 'chín', 'chin',
                    'slide'],
        'description': 'Chuyển đến slide theo số (1-9)'
    },
    {
        'action': 'start_slideshow',
        'keywords': ['trình chiếu', 'trinh chieu', 'bắt đầu', 'bat dau', 'chạy', 'chay'],
        'description': 'Bắt đầu trình chiếu'
    },
    {
        'action': 'next_slide',
        'keywords': ['tiếp tục', 'tiep tuc', 'tiếp', 'tiep', 'kế tiếp', 'ke tiep', 'next'],
        'description': 'Chuyển sang slide tiếp theo'
    },
    {
        'action': 'previous_slide',
        'keywords': ['lùi lại', 'lui lai', 'lùi', 'lui', 'quay lại', 'quay lai', 'trước', 'truoc', 'back'],
        'description': 'Quay lại slide trước'
    },
    {
        'action': 'exit',
        'keywords': ['thoát', 'thoat', 'dừng', 'dung', 'kết thúc', 'ket thuc', 'exit', 'quit'],
        'description': 'Thoát chương trình'
    }
]

# Cấu hình logging
LOGGING_CONFIG = {
    'log_file': 'app.log',
    'log_level': 'INFO',
    'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
}
