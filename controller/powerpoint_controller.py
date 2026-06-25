"""
PowerPoint Controller - Điều khiển ứng dụng PowerPoint
Sử dụng pywin32 để tương tác với PowerPoint trên Windows
"""

import os
import time
from pathlib import Path
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

from config.settings import POWERPOINT_CONFIG
from utils.logger import get_logger

class PowerPointController:
    """Class điều khiển PowerPoint qua COM interface"""
    
    def __init__(self):
        """Khởi tạo controller"""
        self.powerpoint = None
        self.presentation = None
        self.logger = get_logger()
        
        if not WIN32_AVAILABLE:
            self.logger.warning("pywin32 not available - some features may not work")
    
    def execute(self, command):
        """
        Thực thi lệnh điều khiển PowerPoint
        
        Args:
            command (dict): Dictionary chứa action và params
        
        Returns:
            dict: Kết quả thực thi với success và message
        """
        action = command['action']
        params = command.get('params', {})
        
        try:
            # Mapping action với method tương ứng
            action_map = {
                'open_powerpoint': self._open_powerpoint,
                'find_file': self._find_file,
                'goto_slide': self._goto_slide,
                'start_slideshow': self._start_slideshow,
                'next_slide': self._next_slide,
                'previous_slide': self._previous_slide,
                'exit': self._exit
            }
            
            if action in action_map:
                return action_map[action](params)
            else:
                return {
                    'success': False,
                    'message': f"Lệnh không được hỗ trợ: {action}"
                }
        
        except Exception as e:
            self.logger.error(f"Error executing {action}: {str(e)}")
            return {
                'success': False,
                'message': f"Lỗi khi thực thi: {str(e)}"
            }
    
    def _open_powerpoint(self, params):
        """Mở ứng dụng PowerPoint"""
        if not WIN32_AVAILABLE:
            return {'success': False, 'message': 'pywin32 không được cài đặt'}
        
        try:
            self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            self.powerpoint.Visible = True
            return {'success': True, 'message': 'Đã mở PowerPoint'}
        except Exception as e:
            return {'success': False, 'message': f'Không thể mở PowerPoint: {str(e)}'}
    
    def _find_file(self, params):
        """Tìm và mở file PowerPoint trong thư mục demo"""
        if not self.powerpoint:
            self._open_powerpoint({})
        
        # Tìm file .pptx trong thư mục hiện tại và thư mục demo
        search_paths = [
            Path('.'),
            Path('./demo'),
            Path('./presentations'),
            Path(POWERPOINT_CONFIG['default_path'])
        ]
        
        ppt_file = None
        for search_path in search_paths:
            if search_path.exists():
                ppt_files = list(search_path.glob('*.pptx'))
                if ppt_files:
                    ppt_file = ppt_files[0]
                    break
        
        if ppt_file and ppt_file.exists():
            try:
                self.presentation = self.powerpoint.Presentations.Open(str(ppt_file.absolute()))
                return {'success': True, 'message': f'Đã mở file: {ppt_file.name}'}
            except Exception as e:
                return {'success': False, 'message': f'Lỗi mở file: {str(e)}'}
        else:
            return {'success': False, 'message': 'Không tìm thấy file PowerPoint'}
    
    def _goto_slide(self, params):
        """Chuyển đến slide theo số"""
        slide_number = params.get('slide_number', 1)
        
        if not self.presentation:
            return {'success': False, 'message': 'Chưa mở file PowerPoint'}
        
        try:
            total_slides = self.presentation.Slides.Count
            if 1 <= slide_number <= total_slides:
                # Nếu đang trong slideshow mode
                if self.presentation.SlideShowWindow:
                    self.presentation.SlideShowWindow.View.GotoSlide(slide_number)
                else:
                    # Nếu đang ở edit mode
                    self.presentation.Windows(1).View.GotoSlide(slide_number)
                return {'success': True, 'message': f'Đã chuyển đến slide {slide_number}'}
            else:
                return {'success': False, 'message': f'Slide {slide_number} không tồn tại (có {total_slides} slides)'}
        except Exception as e:
            return {'success': False, 'message': f'Lỗi chuyển slide: {str(e)}'}
    
    def _start_slideshow(self, params):
        """Bắt đầu trình chiếu"""
        if not self.presentation:
            return {'success': False, 'message': 'Chưa mở file PowerPoint'}
        
        try:
            self.presentation.SlideShowSettings.Run()
            return {'success': True, 'message': 'Đã bắt đầu trình chiếu'}
        except Exception as e:
            return {'success': False, 'message': f'Lỗi bắt đầu trình chiếu: {str(e)}'}
    
    def _next_slide(self, params):
        """Chuyển sang slide tiếp theo"""
        if not self.presentation:
            return {'success': False, 'message': 'Chưa mở file PowerPoint'}
        
        try:
            if self.presentation.SlideShowWindow:
                self.presentation.SlideShowWindow.View.Next()
                return {'success': True, 'message': 'Đã chuyển slide tiếp theo'}
            else:
                return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
        except Exception as e:
            return {'success': False, 'message': f'Lỗi chuyển slide: {str(e)}'}
    
    def _previous_slide(self, params):
        """Quay lại slide trước"""
        if not self.presentation:
            return {'success': False, 'message': 'Chưa mở file PowerPoint'}
        
        try:
            if self.presentation.SlideShowWindow:
                self.presentation.SlideShowWindow.View.Previous()
                return {'success': True, 'message': 'Đã lùi lại slide trước'}
            else:
                return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
        except Exception as e:
            return {'success': False, 'message': f'Lỗi lùi slide: {str(e)}'}
    
    def _exit(self, params):
        """Thoát chương trình"""
        try:
            if self.presentation:
                self.presentation.Close()
            if self.powerpoint:
                self.powerpoint.Quit()
        except:
            pass
        return {'success': True, 'message': 'Đang thoát...'}
