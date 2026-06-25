"""
Script test microphone và speech recognition
Chạy file này để kiểm tra microphone hoạt động tốt không
"""

import speech_recognition as sr

def test_microphone():
    """Test microphone và Google Speech API"""
    print("="*60)
    print("TEST MICROPHONE & SPEECH RECOGNITION")
    print("="*60)
    
    recognizer = sr.Recognizer()
    
    # Liệt kê các microphone có sẵn
    print("\n📋 Danh sách Microphone:")
    for index, name in enumerate(sr.Microphone.list_microphone_names()):
        print(f"   {index}: {name}")
    
    print("\n🎙️  Đang khởi tạo microphone...")
    
    try:
        with sr.Microphone() as source:
            print("✅ Microphone OK!")
            
            print("\n🔧 Đang điều chỉnh nhiễu nền...")
            recognizer.adjust_for_ambient_noise(source, duration=2)
            print("✅ Đã điều chỉnh xong!")
            
            print("\n🎤 HÃY NÓI GÌ ĐÓ (tiếng Việt)...")
            print("   (Timeout: 5 giây)")
            
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
            print("✅ Đã ghi âm xong!")
            
            print("\n🔄 Đang nhận dạng giọng nói...")
            text = recognizer.recognize_google(audio, language='vi-VN')
            
            print("\n" + "="*60)
            print(f"📝 KẾT QUẢ: '{text}'")
            print("="*60)
            print("\n✅ TEST THÀNH CÔNG! Microphone hoạt động tốt!")
            
    except sr.WaitTimeoutError:
        print("\n❌ TIMEOUT: Không nghe thấy giọng nói trong 5 giây")
        print("   → Kiểm tra microphone có được bật không")
        
    except sr.UnknownValueError:
        print("\n❌ KHÔNG NHẬN DẠNG ĐƯỢC giọng nói")
        print("   → Thử nói to và rõ hơn")
        print("   → Kiểm tra kết nối internet")
        
    except sr.RequestError as e:
        print(f"\n❌ LỖI DỊCH VỤ GOOGLE API: {e}")
        print("   → Kiểm tra kết nối internet")
        
    except Exception as e:
        print(f"\n❌ LỖI: {e}")

if __name__ == "__main__":
    test_microphone()
