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
        
        print("\n" + "=" * 60)
        print("🎤 SPEECH POWERPOINT CONTROLLER")
        print("=" * 60)

        print("\n📌 QUY TRÌNH SỬ DỤNG")

        print("\n【Khởi động】")
        print("  • 'mở powerpoint'")
        print("    → Mở ứng dụng Microsoft PowerPoint.")

        print("\n【Mở bài thuyết trình】")
        print("  • 'tìm file'")
        print("    → Hiển thị danh sách file PowerPoint trong thư mục presentations.")
        print("  • 'số 1', 'số 2', ...")
        print("    → Chọn file theo số thứ tự.")
        print("  • 'mở file'")
        print("    → Mở file PowerPoint đã chọn.")

        print("\n【Điều khiển trình chiếu】")
        print("  • 'trình chiếu'")
        print("    → Bắt đầu chế độ Slide Show.")
        print("  • 'slide 5'")
        print("    → Chuyển đến slide số 5.")
        print("  • 'tiếp tục'")
        print("    → Chuyển sang slide tiếp theo.")
        print("  • 'lùi lại'")
        print("    → Quay về slide trước.")
        print("  • 'thoát'")
        print("    → Thoát chế độ trình chiếu.")

        print("\n【Kết thúc】")
        print("  • 'tắt chương trình'")
        print("    → Đóng PowerPoint và kết thúc ứng dụng.")

        print("\n📌 LƯU Ý")
        print("  • Sau 'tìm file' hãy nói: 'số <n>' để chọn file.")
        print("  • Sau khi chọn file, hãy nói: 'mở file'.")
        print("  • Sau khi mở file, hãy nói: 'trình chiếu'.")
        print("  • Muốn chuyển slide, hãy nói: 'slide <n>'.")
        print("  • 'số <n>' dùng để CHỌN FILE.")
        print("  • 'slide <n>' dùng để CHUYỂN SLIDE.")
        print("  • Muốn kết thúc Slide Show hãy nói: 'thoát'.")
        print("  • Muốn đóng hoàn toàn chương trình hãy nói: 'tắt chương trình'.")

        print("\n" + "=" * 60)
        print("🎯 Sẵn sàng nhận lệnh giọng nói...")
        print("=" * 60 + "\n")
        
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
                    if command['action'] == 'close_program':
                        ppt_controller.execute({'action': 'exit', 'params': {}})
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
