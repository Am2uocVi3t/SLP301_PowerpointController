"""
Speech PowerPoint Controller - Main Entry Point
Điểm khởi chạy chính của ứng dụng điều khiển PowerPoint bằng giọng nói
"""

import sys
from speech.recognizer import SpeechRecognizer
from parser.command_parser import CommandParser
from controller.powerpoint_controller import PowerPointController
from utils.logger import setup_logger

def main():
    """Hàm main để chạy ứng dụng"""
    # Khởi tạo logger
    logger = setup_logger()
    logger.info("=== Speech PowerPoint Controller Starting ===")
    
    try:
        # Khởi tạo các module
        speech_recognizer = SpeechRecognizer()
        command_parser = CommandParser()
        ppt_controller = PowerPointController()
        
        print("\n" + "="*60)
        print("🎤 SPEECH POWERPOINT CONTROLLER")
        print("="*60)
        print("\nCác lệnh hỗ trợ:")
        print("  - 'mở powerpoint' : Mở ứng dụng PowerPoint")
        print("  - 'tìm file'      : Tìm và mở file PowerPoint")
        print("  - 'một/hai/ba'    : Chuyển đến slide 1/2/3")
        print("  - 'trình chiếu'   : Bắt đầu trình chiếu")
        print("  - 'tiếp tục'      : Chuyển slide tiếp theo")
        print("  - 'lùi lại'       : Quay lại slide trước")
        print("  - 'thoát'         : Thoát chương trình")
        print("\n" + "="*60)
        print("Sẵn sàng nhận lệnh giọng nói...")
        print("="*60 + "\n")
        
        # Vòng lặp chính
        while True:
            # Nhận giọng nói từ microphone
            text = speech_recognizer.listen()
            
            if text:
                print(f"\n📝 Đã nhận: '{text}'")
                logger.info(f"Recognized text: {text}")
                
                # Phân tích lệnh
                command = command_parser.parse(text)
                
                if command:
                    print(f"✅ Lệnh: {command['action']}")
                    logger.info(f"Parsed command: {command}")
                    
                    # Thực thi lệnh
                    result = ppt_controller.execute(command)
                    
                    if result['success']:
                        print(f"✔️  {result['message']}")
                    else:
                        print(f"❌ {result['message']}")
                    
                    # Kiểm tra lệnh thoát
                    if command['action'] == 'exit':
                        break
                else:
                    print("❓ Không nhận dạng được lệnh")
            
            print("-" * 60)
    
    except KeyboardInterrupt:
        print("\n\n⚠️  Người dùng dừng chương trình (Ctrl+C)")
        logger.info("User interrupted with Ctrl+C")
    
    except Exception as e:
        print(f"\n❌ Lỗi: {str(e)}")
        logger.error(f"Error in main: {str(e)}", exc_info=True)
    
    finally:
        print("\n👋 Đã thoát Speech PowerPoint Controller")
        logger.info("=== Speech PowerPoint Controller Stopped ===")

if __name__ == "__main__":
    main()
