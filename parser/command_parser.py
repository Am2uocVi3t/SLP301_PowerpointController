"""
Command Parser - Phân tích lệnh từ text
Chuyển đổi text thành các lệnh có cấu trúc để thực thi
"""

from config.settings import COMMANDS
from utils.logger import get_logger

class CommandParser:
    """Class phân tích và xác định lệnh từ text"""
    
    def __init__(self):
        """Khởi tạo parser với danh sách lệnh"""
        self.commands = COMMANDS
        self.logger = get_logger()
    
    def parse(self, text):
        """
        Phân tích text để xác định lệnh
        
        Args:
            text (str): Text cần phân tích
        
        Returns:
            dict: Dictionary chứa action và params, hoặc None nếu không khớp
        """
        if not text:
            return None
        
        text = text.lower().strip()
        
        # Kiểm tra từng lệnh trong config
        for command_config in self.commands:
            action = command_config['action']
            keywords = command_config['keywords']
            
            # Kiểm tra xem có keyword nào trong text không
            if any(keyword in text for keyword in keywords):
                self.logger.info(f"Matched command: {action}")
                
                # Xử lý đặc biệt cho lệnh chuyển slide theo số
                if action == 'goto_slide':
                    slide_number = self._extract_slide_number(text)
                    return {
                        'action': action,
                        'params': {'slide_number': slide_number}
                    }
                
                # Các lệnh khác không cần params
                return {
                    'action': action,
                    'params': {}
                }
        
        # Không tìm thấy lệnh khớp
        self.logger.warning(f"No command matched for text: {text}")
        return None
    
    def _extract_slide_number(self, text):
        """
        Trích xuất số slide từ text
        
        Args:
            text (str): Text chứa số slide
        
        Returns:
            int: Số slide (1-9), mặc định là 1
        """
        # Mapping từ chữ sang số
        number_mapping = {
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
        
        # Tìm số trong text
        for word, number in number_mapping.items():
            if word in text:
                return number
        
        return 1  # Mặc định slide 1
