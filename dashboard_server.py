"""
Dashboard Server - Flask + SocketIO backend cho Dashboard
Tích hợp toàn bộ Speech Recognition loop và emit real-time events
"""

import os
import sys
import time
import threading
from pathlib import Path
from datetime import datetime

from flask import Flask, send_from_directory
from flask_socketio import SocketIO

# Đảm bảo import được các module của project
sys.path.insert(0, str(Path(__file__).resolve().parent))

from utils.logger import setup_logger

# ============================================================
# Flask App Setup
# ============================================================

app = Flask(__name__, static_folder=None)
app.config['SECRET_KEY'] = 'slp301-dashboard-secret'
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# Logger
logger = setup_logger()

# Dashboard directory
DASHBOARD_DIR = Path(__file__).resolve().parent / 'dashboard'

# ============================================================
# State Management
# ============================================================

class DashboardState:
    """Quản lý trạng thái hệ thống cho Dashboard"""

    def __init__(self):
        self.mic_status = 'idle'          # idle | listening | processing | error
        self.ppt_status = 'disconnected'  # disconnected | connected | file_opened | slideshow
        self.ppt_detail = ''
        self.current_slide = 0
        self.total_slides = 0
        self.command_history = []         # List of command entries
        self.running = False

    def add_command(self, raw_text, action, success, message):
        """Thêm lệnh vào lịch sử"""
        entry = {
            'raw_text': raw_text,
            'action': action,
            'success': success,
            'message': message,
            'timestamp': datetime.now().strftime('%H:%M:%S')
        }
        self.command_history.insert(0, entry)
        # Giới hạn 50 entries
        if len(self.command_history) > 50:
            self.command_history = self.command_history[:50]
        return entry


state = DashboardState()

# ============================================================
# Routes
# ============================================================

@app.route('/')
def index():
    """Serve trang Dashboard"""
    return send_from_directory(str(DASHBOARD_DIR), 'index.html')

@app.route('/style.css')
def style():
    """Serve CSS"""
    return send_from_directory(str(DASHBOARD_DIR), 'style.css')

# ============================================================
# SocketIO Events
# ============================================================

@socketio.on('connect')
def handle_connect():
    """Khi Dashboard kết nối"""
    logger.info('Dashboard client connected')
    emit_log('info', 'Dashboard client connected')

    # Gửi trạng thái hiện tại
    socketio.emit('mic_status', {'status': state.mic_status})
    socketio.emit('ppt_status', {
        'status': state.ppt_status,
        'detail': state.ppt_detail
    })
    socketio.emit('current_slide', {
        'current': state.current_slide,
        'total': state.total_slides
    })

    # Gửi command history
    if state.command_history:
        socketio.emit('command_history', {'entries': state.command_history})


@socketio.on('disconnect')
def handle_disconnect():
    """Khi Dashboard ngắt kết nối"""
    logger.info('Dashboard client disconnected')

# ============================================================
# Emit Helpers
# ============================================================

def emit_mic_status(status, message=''):
    """Emit trạng thái microphone"""
    state.mic_status = status
    socketio.emit('mic_status', {'status': status, 'message': message})


def emit_ppt_status(status, detail=''):
    """Emit trạng thái PowerPoint"""
    state.ppt_status = status
    state.ppt_detail = detail
    socketio.emit('ppt_status', {'status': status, 'detail': detail})


def emit_current_command(raw_text, action, success, message):
    """Emit lệnh hiện tại + thêm vào history"""
    entry = state.add_command(raw_text, action, success, message)
    socketio.emit('current_command', entry)


def emit_slide_info(current, total):
    """Emit thông tin slide"""
    state.current_slide = current
    state.total_slides = total
    socketio.emit('current_slide', {'current': current, 'total': total})


def emit_log(level, message):
    """Emit log entry"""
    socketio.emit('log_entry', {
        'level': level,
        'message': message,
        'timestamp': datetime.now().strftime('%H:%M:%S')
    })

# ============================================================
# Speech Recognition Loop (chạy trên thread riêng)
# ============================================================

def update_ppt_state(ppt_controller):
    """Cập nhật trạng thái PowerPoint từ controller"""
    try:
        if ppt_controller.waiting_for_file_selection:
            emit_ppt_status('waiting', f'Đã tìm thấy {len(ppt_controller.available_files)} file. Vui lòng xem Log.')
            emit_slide_info(0, 0)
            return

        if ppt_controller.powerpoint is None:
            emit_ppt_status('disconnected')
            return

        if ppt_controller.presentation is not None:
            file_name = ''
            try:
                file_name = ppt_controller.presentation.Name
            except Exception:
                if ppt_controller.selected_file:
                    file_name = ppt_controller.selected_file.name

            if ppt_controller.slideshow_running:
                emit_ppt_status('slideshow', file_name)
                # Cập nhật slide info
                try:
                    window = ppt_controller.powerpoint.SlideShowWindows(1)
                    current = window.View.Slide.SlideIndex
                    total = ppt_controller.presentation.Slides.Count
                    emit_slide_info(current, total)
                except Exception:
                    pass
            else:
                emit_ppt_status('file_opened', file_name)
                try:
                    total = ppt_controller.presentation.Slides.Count
                    emit_slide_info(0, total)
                except Exception:
                    pass
        else:
            emit_ppt_status('connected')
            emit_slide_info(0, 0)
    except Exception:
        emit_ppt_status('disconnected')
        emit_slide_info(0, 0)


def speech_loop():
    """Vòng lặp chính cho speech recognition"""
    logger.info("=== Speech PowerPoint Controller Starting (Dashboard Mode) ===")
    emit_log('info', '=== Speech PowerPoint Controller Starting ===')

    try:
        # Import lazy — để dashboard vẫn chạy được nếu thiếu thư viện
        from speech.recognizer import SpeechRecognizer
        from parser.command_parser import CommandParser
        from controller.powerpoint_controller import PowerPointController

        # Khởi tạo các module
        emit_log('info', 'Đang khởi tạo SpeechRecognizer...')
        emit_mic_status('processing', 'Đang khởi tạo...')
        speech_recognizer = SpeechRecognizer()
        emit_log('info', 'SpeechRecognizer sẵn sàng')

        emit_log('info', 'Đang khởi tạo CommandParser...')
        command_parser = CommandParser()
        emit_log('info', 'CommandParser sẵn sàng')

        emit_log('info', 'Đang khởi tạo PowerPointController...')
        ppt_controller = PowerPointController()
        emit_log('info', 'PowerPointController sẵn sàng')

        emit_mic_status('idle')
        emit_log('info', '🎯 Hệ thống sẵn sàng nhận lệnh giọng nói!')

        state.running = True

        while state.running:
            # Update PPT state
            update_ppt_state(ppt_controller)

            # Listening
            emit_mic_status('listening')
            emit_log('info', 'Đang lắng nghe...')

            text = speech_recognizer.listen()

            if text:
                emit_mic_status('processing')
                emit_log('info', f'Đã nhận: "{text}"')
                logger.info(f"Recognized text: {text}")

                # Parse command
                command = command_parser.parse(text)

                if command:
                    action = command['action']
                    emit_log('info', f'Lệnh: {action}')
                    logger.info(f"Parsed command: {command}")

                    # Execute
                    result = ppt_controller.execute(command)
                    success = result['success']
                    message = result['message']

                    short_message = message.split('\n')[0] if '\n' in message else message
                    emit_current_command(text, action, success, short_message)

                    if success:
                        emit_log('info', f'✅ {message}')
                    else:
                        emit_log('warning', f'❌ {message}')

                    # Update PPT state sau khi execute
                    update_ppt_state(ppt_controller)

                    # Kiểm tra lệnh thoát
                    if action == 'close_program':
                        emit_log('info', 'Đang thoát chương trình...')
                        ppt_controller.execute({'action': 'exit', 'params': {}})
                        state.running = False
                        break
                else:
                    emit_current_command(text, 'unknown', False, 'Không nhận dạng được lệnh')
                    emit_log('warning', f'Không nhận dạng được lệnh: "{text}"')
            else:
                emit_mic_status('idle')

    except KeyboardInterrupt:
        emit_log('warning', 'Người dùng dừng chương trình (Ctrl+C)')
        logger.info("User interrupted with Ctrl+C")

    except Exception as e:
        emit_log('error', f'Lỗi: {str(e)}')
        logger.error(f"Error in speech_loop: {str(e)}", exc_info=True)

    finally:
        state.running = False
        emit_mic_status('idle')
        emit_log('info', '=== Speech PowerPoint Controller Stopped ===')
        logger.info("=== Speech PowerPoint Controller Stopped ===")

# ============================================================
# Main
# ============================================================

def main():
    """Khởi chạy Dashboard Server"""
    # Fix encoding cho Windows console
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

    print("\n" + "=" * 60)
    print("  SPEECH POWERPOINT CONTROLLER - DASHBOARD MODE")
    print("=" * 60)
    print(f"\n  Dashboard URL: http://localhost:5000")
    print(f"  Dashboard Dir: {DASHBOARD_DIR}")
    print("\n  Mo trinh duyet tai http://localhost:5000 de xem Dashboard")
    print("  Nhan Ctrl+C de dung server")
    print("=" * 60 + "\n")

    # Chạy speech loop trên thread riêng
    speech_thread = threading.Thread(target=speech_loop, daemon=True)
    speech_thread.start()

    # Chạy Flask server
    try:
        socketio.run(app, host='0.0.0.0', port=5000, debug=False, allow_unsafe_werkzeug=True)
    except KeyboardInterrupt:
        print("\n\n  Server dang dung...")
        state.running = False
    finally:
        print("  Da tat Dashboard Server")


if __name__ == '__main__':
    main()
