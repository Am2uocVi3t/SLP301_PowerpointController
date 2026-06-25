"""
Speech Recognizer - Module nhận dạng giọng nói
Sử dụng thư viện SpeechRecognition để chuyển giọng nói thành text
"""

import speech_recognition as sr
from config.settings import SPEECH_CONFIG
from utils.logger import get_logger

class SpeechRecognizer:
    """Class xử lý nhận dạng giọng nói từ microphone"""
    
    def __init__(self):
        """Khởi tạo recognizer và microphone"""
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.logger = get_logger()
        
        # Điều chỉnh noise và energy threshold
        with self.microphone as source:
            print("🎙️  Đang hiệu chỉnh microphone cho nhiễu nền...")
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            print("✅ Microphone sẵn sàng!")
    
    def listen(self):
        """
        Lắng nghe và nhận dạng giọng nói
        
        Returns:
            str: Text được nhận dạng, hoặc None nếu không nhận dạng được
        """
        try:
            with self.microphone as source:
                print("\n🎤 Đang lắng nghe... (hãy nói lệnh)")
                
                # Ghi âm từ microphone
                audio = self.recognizer.listen(
                    source,
                    timeout=SPEECH_CONFIG['timeout'],
                    phrase_time_limit=SPEECH_CONFIG['phrase_time_limit']
                )
                
                print("🔄 Đang xử lý...")
                
                # Nhận dạng giọng nói bằng Google Speech Recognition
                text = self.recognizer.recognize_google(
                    audio,
                    language=SPEECH_CONFIG['language']
                )
                
                return text.lower()  # Chuyển về chữ thường để dễ xử lý
        
        except sr.WaitTimeoutError:
            print("⏰ Timeout - Không nghe thấy giọng nói")
            return None
        
        except sr.UnknownValueError:
            print("❓ Không nhận dạng được giọng nói")
            self.logger.warning("Could not understand audio")
            return None
        
        except sr.RequestError as e:
            print(f"❌ Lỗi dịch vụ nhận dạng: {e}")
            self.logger.error(f"Recognition service error: {e}")
            return None
        
        except Exception as e:
            print(f"❌ Lỗi không xác định: {e}")
            self.logger.error(f"Unexpected error in listen: {e}")
            return None
