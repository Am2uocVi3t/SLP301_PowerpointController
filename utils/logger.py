"""
Logger - Hệ thống ghi log
Ghi lại các hoạt động và lỗi của ứng dụng
"""

import logging
from config.settings import LOGGING_CONFIG

# Logger singleton
_logger = None

def setup_logger():
    """
    Thiết lập logger cho toàn bộ ứng dụng
    
    Returns:
        logging.Logger: Logger đã được cấu hình
    """
    global _logger
    
    if _logger is not None:
        return _logger
    
    # Tạo logger
    _logger = logging.getLogger('SpeechPPTController')
    _logger.setLevel(getattr(logging, LOGGING_CONFIG['log_level']))
    
    # Tạo file handler
    file_handler = logging.FileHandler(
        LOGGING_CONFIG['log_file'],
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    
    # Tạo console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.ERROR)
    
    # Tạo formatter
    formatter = logging.Formatter(LOGGING_CONFIG['format'])
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Thêm handlers
    _logger.addHandler(file_handler)
    _logger.addHandler(console_handler)
    
    return _logger

def get_logger():
    """
    Lấy logger đã được thiết lập
    
    Returns:
        logging.Logger: Logger instance
    """
    global _logger
    if _logger is None:
        return setup_logger()
    return _logger
