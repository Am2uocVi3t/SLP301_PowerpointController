"""
Command Parser - Phân tích lệnh từ text
Chuyển đổi text thành các lệnh có cấu trúc để thực thi
"""
import re
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

        # ==========================================================
        # 1. Ưu tiên chọn FILE
        # Cú pháp:
        #   "số 1"
        #   "số 2"
        # ==========================================================

        file_index = self._extract_file_number(text)

        if file_index is not None:
            self.logger.info(f"Matched command: select_file with index {file_index}")
            return {
                'action': 'select_file',
                'params': {'file_index': file_index}
            }
        
        # ==========================================================
        # 2. Ưu tiên chuyển SLIDE
        # Cú pháp:
        #   "slide 5"
        #   "slide năm"
        #   "trang 3"
        # ==========================================================

        slide_number = self._extract_slide_number(text)
        
        if slide_number is not None:
            self.logger.info(f"Matched command: goto_slide with number {slide_number}")
            return {
                'action': 'goto_slide',
                'params': {'slide_number': slide_number}
            }

        # ==========================================================
        # 3. Các command thông thường
        # ==========================================================

        for command_config in self.commands:
            action = command_config['action']

            #Bỏ qua 2 command đã xử lý
            if action in ['select_file', 'goto_slide']:
                continue

            keywords = command_config['keywords']

            if any(keyword in text for keyword in keywords):
                self.logger.info(f"Matched command: {action}")
                return {
                    'action': action,
                    'params': {}
                }
            
        self.logger.warning(f"No command matched for text: {text}")
        return None


    def _extract_file_number(self, text):
        """
        Nhận:
            số 1
            số 25
            số ba
        """

        match = re.search(r"^(số|so)\s+(.+)$", text)

        if not match:
            return None

        value = match.group(2).strip()

        return self._extract_number(value)
    

    def _extract_slide_number(self, text):
        """
        Nhận:
            slide 1
            slide 10
            slide 35

            slide ba
            trang 5
        """

        match = re.search(r"^(slide|trang)\s+(.+)$", text)

        if not match:
            return None

        value = match.group(2).strip()

        return self._extract_number(value)
    

    def _extract_number(self, value):
        """
        Chuyển text thành số.

        Hỗ trợ:
            1
            10
            25
            100

            một
            hai
            ba
            ...
            chín
        """

        value = value.strip().lower()

        # ==================================================
        # Nếu là số (10, 25, 100...)
        # ==================================================

        if value.isdigit():
            return int(value)

        # ==================================================
        # Nếu là số bằng chữ
        # ==================================================

        number_mapping = {
            'một': 1,
            'mot': 1,

            'hai': 2,

            'ba': 3,

            'bốn': 4,
            'bon': 4,
            'tư': 4,
            'tu': 4,

            'năm': 5,
            'nam': 5,

            'sáu': 6,
            'sau': 6,

            'bảy': 7,
            'bay': 7,
            'bẩy': 7,

            'tám': 8,
            'tam': 8,

            'chín': 9,
            'chin': 9,
        }

        return number_mapping.get(value)