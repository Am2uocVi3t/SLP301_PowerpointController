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
        'keywords': ['tìm file', 'tim file', 'tìm kiếm file'],
        'description': 'Tìm và mở file PowerPoint'
    },
    {
        'action': 'select_file',
        'keywords': ['số', 'so','một', 'mot', 'hai', 'ba', 'bốn', 'bon', 'tư', 'tu', 'năm', 'nam', 
                    'sáu', 'sau', 'bảy', 'bay', 'tám', 'tam', 'chín', 'chin'],
        'description': 'Chọn slide theo số (1-9)'
    },
    {
        'action': 'open_selected_file',
        'keywords': ['mở file', 'mo file', 'mở tệp'],
        'description': 'Mở file PowerPoint đã chọn'
    },
    {
        'action': 'goto_slide',
        'keywords': [
            'slide',
            'trang',
            'đến slide',
            'toi slide',
            'chuyển slide'
        ],
        'description': 'Chuyển tới slide'
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
        'action': 'exit_slideshow',
        'keywords': ['thoát', 'thoat', 'dừng', 'dung', 'kết thúc', 'ket thuc'],
        'description': 'Thoát trình chiếu và đóng PowerPoint'
    },
    {
        'action': 'close_program',
        'keywords': ['tắt', 'tat', 'tắt chương trình', 'tat chuong trinh', 'đóng', 'dong', 'đóng chương trình', 'dong chuong trinh', 'thoát chương trình', 'thoat chuong trinh'],
        'description': 'Thoát chương trình'
    }
]

# Cấu hình logging
LOGGING_CONFIG = {
    'log_file': 'app.log',
    'log_level': 'INFO',
    'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
}

NUMBER_WORDS = {
    'một': 1, 'mot': 1, '1': 1,
    'hai': 2, '2': 2,
    'ba': 3, '3': 3,
    'bốn': 4, 'bon': 4, 'tư': 4, 'tu': 4, '4': 4,
    'năm': 5, 'nam': 5, '5': 5,
    'sáu': 6, 'sau': 6, '6': 6,
    'bảy': 7, 'bay': 7, 'bẩy': 7, '7': 7,
    'tám': 8, 'tam': 8, '8': 8,
    'chín': 9, 'chin': 9, '9': 9
}