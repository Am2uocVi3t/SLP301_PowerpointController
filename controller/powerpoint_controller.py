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
        
        self.project_root = Path(__file__).resolve().parent.parent
        self.logger.info(f"Root path: {self.project_root}")

        self.available_files = [] #Danh sách file trong presentations
        self.selected_file = None #File được chọn để mở
        self.waiting_for_file_selection  = False #Chờ người dùng chọn file
        self.slideshow_running = False #Trạng thái trình chiếu

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
                'select_file': self._select_file,
                'open_selected_file': self._open_selected_file,
                'goto_slide': self._goto_slide,
                'start_slideshow': self._start_slideshow,
                'next_slide': self._next_slide,
                'previous_slide': self._previous_slide,
                'exit_slideshow': self._exit_slideshow,
                'close_program': self._close_program,
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
        if self.powerpoint is not None:
            try:
                _ = self.powerpoint.Presentations.Count  # Kiểm tra xem PowerPoint còn chạy không
                return {
                    'success': True,
                    'message': 'PowerPoint đã mở'
                }
            except Exception:
                self.powerpoint = None  # Nếu PowerPoint đã bị đóng, reset lại

        if not WIN32_AVAILABLE:
            return {
                'success': False,
                'message': 'pywin32 không được cài đặt. Không thể mở PowerPoint.'
        }
        
        try:
            self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            self.powerpoint.Visible = True
            return {
                'success': True,
                'message': 'Đã mở PowerPoint'
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'Không thể mở PowerPoint: {str(e)}'
            }

    
    def _find_file(self, params):
        """Quét  thư mục presentations và hiển thị danh sách files"""
        self.presentation = None
        self.selected_file = None
        self.waiting_for_file_selection  = True

        presentations_dir = self.project_root / "presentations"

        if not presentations_dir.exists():
            return {
                'success': False,
                'message': f'Thư mục presentations không tồn tại: {presentations_dir}'
            }

        self.available_files = sorted(
            list(presentations_dir.glob('*.pptx')) + list(presentations_dir.glob('*.ppt')),
        )

        if len(self.available_files) == 0:
            return {
                'success': False,
                'message': f'Không tìm thấy file PowerPoint trong thư mục: {presentations_dir}'
            }
        
        self.selected_file = None

        print("\n" + "=" * 60)
        print("DANH SÁCH FILE POWERPOINT")
        print("=" * 60)

        for i, file in enumerate(self.available_files, start=1):
            print(f"{i}. {file.name}")

        print("=" * 60)

        return {
            "success": True,
            "message": "Hãy nói số thứ tự để chọn file."
        }
    

    def _select_file(self, params):
        """Chọn file PowerPoint từ danh sách"""
        if not self.available_files:
            return {
                'success': False,
                'message': 'Chưa có danh sách file. Hãy nói "tìm file" trước.'
            }
        
        if not self.waiting_for_file_selection :
            return {
                'success': False,
                'message': 'Hãy nói "tìm file" trước.'
            }
        
        index = params.get('file_index')

        if index is None:
            return {
                'success': False,
                'message': 'Hãy nói: số 1, số 2, số 3... để chọn file tương ứng.'
            }
        try:
            index = int(index)
        except:
            return {
                'success': False,
                'message': 'Không đọc được số file'
            }

        if index < 1 or index > len(self.available_files):
            return {
                'success': False,
                'message': f'Số thứ tự không hợp lệ. Hãy chọn từ 1 đến {len(self.available_files)}.'
            }
        
        self.selected_file = self.available_files[index - 1]
        self.waiting_for_file_selection  = False

        return {
            'success': True,
            'message': f'Bạn đã chọn file: {self.selected_file.name}. Hãy nói "mở file" để mở.'
        }


    def _open_selected_file(self, params):
        """Mở file PowerPoint đã chọn"""
        if not self.powerpoint:
            result = self._open_powerpoint({})
            if not result['success']:
                return result

        if not self.selected_file:
            return {
                'success': False,
                'message': 'Chưa chọn file PowerPoint. Hãy nói "tìm file" để chọn.'
            }
        
        try:
            self.presentation = self.powerpoint.Presentations.Open(str(self.selected_file.resolve()))
            return {
                'success': True,
                'message': f'Đã mở file: {self.selected_file.name}'
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'Lỗi mở file: {str(e)}'
            }

    def _goto_slide(self, params):
        """Chuyển đến slide theo số"""
        slide_number = params.get('slide_number', 1)
        
        if not self.presentation:
            return {
                'success': False, 
                'message': 'Chưa mở file PowerPoint. Hãy nói "tìm file", chọn số rồi nói "mở file".'
                }
        
        try:
            total = self.presentation.Slides.Count

            if slide_number < 1 or slide_number > total:
                
                return {
                    'success': False, 
                    'message': f'Slide {slide_number} không tồn tại.'
                }

            try:
                window = self.powerpoint.SlideShowWindows(1)
                window.View.GotoSlide(slide_number)

            except Exception:
                self.presentation.Windows(1).View.GotoSlide(slide_number)
            
            return {
                'success': True, 
                'message': f'Đã chuyển đến slide {slide_number}'
            }
        except Exception as e:
                return {
                    'success': False, 
                    'message': f'Lỗi chuyển slide: {str(e)}'
                }

    
    def _start_slideshow(self, params):
        """Bắt đầu trình chiếu"""
        if not self.presentation:
            return {
                'success': False, 
                'message': 'Chưa mở file PowerPoint. Hãy nói "tìm file", chọn số rồi nói "mở file".'
                }
        
        try:
            self.presentation.SlideShowSettings.Run()
            time.sleep(0.3) # Đợi một chút để slideshow bắt đầu

            self.slideshow_running = True

            return {'success': True, 'message': 'Đã bắt đầu trình chiếu'}
        
        except Exception as e:
            return {
                'success': False, 
                'message': f'Lỗi bắt đầu trình chiếu: {str(e)}'
                }
    
    def _next_slide(self, params):
        """Chuyển sang slide tiếp theo"""
        
        #Chưa mở file
        if not self.presentation:
            return {
                'success': False, 
                'message': 'Chưa mở file PowerPoint. Hãy nói "tìm file", chọn số rồi nói "mở file".'
                }
        
        #Đã mở file nhưng chưa trình chiếu
        if not self.slideshow_running:
            return {
                    'success': False,
                    'message': 'Bạn chưa nói bắt đầu trình chiếu. Hãy nói "trình chiếu" trước.'
                }
        
        try:
            window = self.powerpoint.SlideShowWindows(1)
            window.View.Next()
            
            return {
                'success': True,
                'message': 'Đã chuyển slide tiếp theo'
            }
        
        except Exception as e:

            return {
                'success': False,
                'message': f'Lỗi chuyển slide: {str(e)}'
            }
        #     if self.presentation.SlideShowWindow:
        #         self.presentation.SlideShowWindow.View.Next()
        #         return {'success': True, 'message': 'Đã chuyển slide tiếp theo'}
        #     else:
        #         return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
        # except Exception as e:
        #     return {'success': False, 'message': f'Lỗi chuyển slide: {str(e)}'}
    
    def _previous_slide(self, params):
        """Quay lại slide trước"""
        #Chưa mở file
        if not self.presentation:
            return {
                'success': False, 
                'message': 'Chưa mở file PowerPoint. Hãy nói "tìm file", chọn số rồi nói "mở file".'
                }
        
        #Đã mở file nhưng chưa trình chiếu
        if not self.slideshow_running:
            return {
                    'success': False,
                    'message': 'Bạn chưa nói bắt đầu trình chiếu. Hãy nói "trình chiếu" trước.'
                }
        
        try:
            window = self.powerpoint.SlideShowWindows(1)
            window.View.Previous()

            return {
                'success': True,
                'message': 'Đã quay lại slide trước'
            }
        
        except Exception as e:
            return {
                'success': False,
                'message': f'Lỗi quay lại slide: {str(e)}'
            }
        #     if self.presentation.SlideShowWindow:
        #         self.presentation.SlideShowWindow.View.Previous()
        #         return {'success': True, 'message': 'Đã lùi lại slide trước'}
        #     else:
        #         return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
        # except Exception as e:
        #     return {'success': False, 'message': f'Lỗi lùi slide: {str(e)}'}
    
    def _exit_slideshow(self, params):
        """Thoát chế độ trình chiếu"""
        if not self.presentation:
            return {
                'success': False,
                'message': 'Chưa mở file PowerPoint. Hãy nói "tìm file", chọn số rồi nói "mở file".'
            }
        
        try:
            window = self.powerpoint.SlideShowWindows(1)
            window.View.Exit()

            return {
                'success': True,
                'message': 'Đã thoát chế độ trình chiếu'
            }
        
        except Exception:
            return {
                'success': False,
                'message': 'Không trong chế độ trình chiếu'
            }


    def _close_program(self, params):
        """Đóng PowerPoint và thoát chương trình"""
        try:
            if self.presentation:
                self.presentation.Close()
                self.presentation = None
            
            if self.powerpoint:
                self.powerpoint.Quit()
                self.powerpoint = None
            
            return {
                'success': True,
                'message': 'Đã đóng PowerPoint và thoát chương trình'
            }
        
        except Exception as e:
            return {
                'success': False,
                'message': f'Lỗi khi đóng PowerPoint: {str(e)}'
            }
