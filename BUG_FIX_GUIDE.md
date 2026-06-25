# 🔧 BUG FIX QUICK REFERENCE

This is a cheat sheet for fixing the 6 critical bugs.

---

## 🔴 BUG-001: Missing Demo File
**Assigned**: Hưng  
**Priority**: P0 - CRITICAL  
**Time**: 1 hour

### What's Wrong
No PowerPoint file exists for demo. The "tìm file" command will fail.

### How to Fix
1. Open PowerPoint
2. Create new presentation
3. Add 7 slides:
   ```
   Slide 1: "Speech PowerPoint Controller Demo"
   Slide 2: "Slide 2 - Navigation Test"
   Slide 3: "Slide 3 - Vietnamese Voice Commands"
   Slide 4: "Slide 4"
   Slide 5: "Slide 5"
   Slide 6: "Slide 6"
   Slide 7: "Slide 7 - End"
   ```
4. Save as: `demo/sample.pptx`
5. Test: run `python main.py` and say "tìm file"

### Verification
```bash
# File should exist
ls demo/sample.pptx

# Should see output
demo/sample.pptx
```

---

## 🔴 BUG-002: SlideShowWindow Check
**Assigned**: Nam  
**Priority**: P0 - CRITICAL  
**Time**: 3 hours

### What's Wrong
Lines 140, 161, 172 in `controller/powerpoint_controller.py`:
```python
if self.presentation.SlideShowWindow:  # ❌ WRONG
```
This checks if the attribute exists, not if slideshow is ACTIVE.

### How to Fix

**Method 1: _goto_slide() - Line 131-151**
```python
def _goto_slide(self, params):
    """Chuyển đến slide theo số"""
    slide_number = params.get('slide_number', 1)
    
    if not self.presentation:
        return {'success': False, 'message': 'Chưa mở file PowerPoint'}
    
    try:
        total_slides = self.presentation.Slides.Count
        if not (1 <= slide_number <= total_slides):
            return {'success': False, 
                   'message': f'Slide {slide_number} không tồn tại (có {total_slides} slides)'}
        
        # Try slideshow mode first
        try:
            slideshow = self.presentation.SlideShowWindow
            if slideshow and hasattr(slideshow, 'View'):
                slideshow.View.GotoSlide(slide_number)
                return {'success': True, 'message': f'Đã chuyển đến slide {slide_number}'}
        except (AttributeError, Exception):
            pass  # Not in slideshow, try edit mode
        
        # Edit mode
        self.presentation.Windows(1).View.GotoSlide(slide_number)
        return {'success': True, 'message': f'Đã chuyển đến slide {slide_number}'}
        
    except Exception as e:
        return {'success': False, 'message': f'Lỗi chuyển slide: {str(e)}'}
```


**Method 2: _next_slide() - Line 159-170**
```python
def _next_slide(self, params):
    """Chuyển sang slide tiếp theo"""
    if not self.presentation:
        return {'success': False, 'message': 'Chưa mở file PowerPoint'}
    
    try:
        slideshow = self.presentation.SlideShowWindow
        if slideshow and hasattr(slideshow, 'View'):
            slideshow.View.Next()
            return {'success': True, 'message': 'Đã chuyển slide tiếp theo'}
        else:
            return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
    except (AttributeError, Exception):
        return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
```

**Method 3: _previous_slide() - Line 172-183**
```python
def _previous_slide(self, params):
    """Quay lại slide trước"""
    if not self.presentation:
        return {'success': False, 'message': 'Chưa mở file PowerPoint'}
    
    try:
        slideshow = self.presentation.SlideShowWindow
        if slideshow and hasattr(slideshow, 'View'):
            slideshow.View.Previous()
            return {'success': True, 'message': 'Đã lùi lại slide trước'}
        else:
            return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
    except (AttributeError, Exception):
        return {'success': False, 'message': 'Không trong chế độ trình chiếu'}
```

### Verification
1. Start program: `python main.py`
2. Say: "mở powerpoint"
3. Say: "tìm file"
4. Say: "trình chiếu"
5. Say: "tiếp tục" ← Should work now!
6. Say: "lùi lại" ← Should work now!

---

## 🔴 BUG-005: Invalid requirements.txt
**Assigned**: Phát  
**Priority**: P0 - CRITICAL  
**Time**: 15 minutes

### What's Wrong
Lines 13-14 have invalid package names:
```
standard-aifc
standard-distutils
```

### How to Fix
Edit `requirements.txt` to:
```txt
# Speech PowerPoint Controller - Requirements

# Core dependencies
SpeechRecognition==3.10.1
PyAudio==0.2.14
pywin32==306
```

That's it! Remove all other lines.

### Verification
```bash
# Create new venv
python -m venv test_env
test_env\Scripts\activate

# Install
pip install -r requirements.txt

# Should succeed without errors
```

---

## 🟡 BUG-003: Keyword "sau" Conflict
**Assigned**: Việt  
**Priority**: P1 - HIGH  
**Time**: 1.5 hours

### What's Wrong
In `config/settings.py`, "sau" appears in TWO places:
- Line 21: `goto_slide` keywords (means "6")
- Line 38: `previous_slide` keywords (means "after")

Parser will always match `goto_slide` first!

### How to Fix
Edit `config/settings.py`:

```python
# Line 18-26: BEFORE
{
    'action': 'goto_slide',
    'keywords': ['một', 'mot', 'hai', 'ba', 'bốn', 'bon', 'tư', 'tu', 'năm', 'nam', 
                'sáu', 'sau', 'bảy', 'bay', 'tám', 'tam', 'chín', 'chin',
                'slide'],
    'description': 'Chuyển đến slide theo số (1-9)'
},

# AFTER (remove 'sau'):
{
    'action': 'goto_slide',
    'keywords': ['một', 'mot', 'hai', 'ba', 'bốn', 'bon', 'tư', 'tu', 'năm', 'nam', 
                'sáu', 'bảy', 'bay', 'tám', 'tam', 'chín', 'chin',
                'slide'],  # Removed 'sau'
    'description': 'Chuyển đến slide theo số (1-9)'
},
```

Keep "sau" in `previous_slide` only.

### Verification
Test both commands:
```
"slide sáu"  → Should go to slide 6 ✅
"sau"        → Should go to slide 6 ✅
"lùi sau"    → Should go back ✅
```

---

## 🟡 BUG-006: Path Handling
**Assigned**: Hưng  
**Priority**: P1 - HIGH  
**Time**: 2 hours

### What's Wrong
Line 115-120 in `controller/powerpoint_controller.py` uses relative paths:
```python
search_paths = [
    Path('.'),           # ❌ Depends on working directory
    Path('./demo'),      # ❌ Depends on working directory
    Path('./presentations'),
    Path(POWERPOINT_CONFIG['default_path'])
]
```

If you run from a different directory, it breaks!

### How to Fix

**Step 1**: Add import at top of file
```python
# At line 8, add:
import os
from pathlib import Path
```

**Step 2**: Add script_dir in __init__
```python
# In __init__ method (around line 20), add:
def __init__(self):
    """Khởi tạo controller"""
    self.powerpoint = None
    self.presentation = None
    self.logger = get_logger()
    
    # ADD THIS: Get absolute path to script directory
    self.script_dir = Path(__file__).parent.parent
    
    if not WIN32_AVAILABLE:
        self.logger.warning("pywin32 not available - some features may not work")
```

**Step 3**: Update _find_file method
```python
# Around line 115, change to:
def _find_file(self, params):
    """Tìm và mở file PowerPoint trong thư mục demo"""
    if not self.powerpoint:
        self._open_powerpoint({})
    
    # Use absolute paths
    search_paths = [
        self.script_dir,
        self.script_dir / 'demo',
        self.script_dir / 'presentations',
    ]
    
    # ... rest of code stays same
```

### Verification
```bash
# Test from different directory
cd ..
python speech-powerpoint-controller/main.py

# Should still find demo file
```

---

## 🟡 BUG-004: COM Exception Handling
**Assigned**: Nam  
**Priority**: P1 - HIGH  
**Time**: 2 hours

### What's Wrong
If PowerPoint COM interface fails, exceptions aren't caught properly.

### How to Fix

Wrap all COM calls in better try-except:

```python
def _open_powerpoint(self, params):
    """Mở ứng dụng PowerPoint"""
    if not WIN32_AVAILABLE:
        return {'success': False, 'message': 'pywin32 không được cài đặt'}
    
    try:
        # ADD: Check if already open
        if self.powerpoint:
            return {'success': True, 'message': 'PowerPoint đã mở sẵn'}
        
        self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        self.powerpoint.Visible = True
        
        # ADD: Small delay to ensure it's ready
        import time
        time.sleep(0.5)
        
        return {'success': True, 'message': 'Đã mở PowerPoint'}
        
    except Exception as e:
        self.logger.error(f"Failed to open PowerPoint: {str(e)}")
        return {'success': False, 'message': f'Không thể mở PowerPoint: {str(e)}'}
```

Also add retry logic:
```python
def _open_powerpoint_with_retry(self, params, retries=2):
    """Mở PowerPoint với retry logic"""
    for attempt in range(retries):
        result = self._open_powerpoint(params)
        if result['success']:
            return result
        
        if attempt < retries - 1:
            print(f"Thử lại lần {attempt + 2}...")
            import time
            time.sleep(1)
    
    return result
```

### Verification
1. Close PowerPoint manually
2. Run program
3. Say "mở powerpoint"
4. Should handle errors gracefully

---

## 📊 FIX PRIORITY ORDER

**Do these IN ORDER**:

1. **BUG-005** (Phát - 15 min) ← Fix installation first!
2. **BUG-001** (Hưng - 1h) ← Demo file needed next
3. **BUG-002** (Nam - 3h) ← Core functionality
4. **Test Everything** (All - 1h)
5. **BUG-006** (Hưng - 2h) ← Path reliability
6. **BUG-003** (Việt - 1.5h) ← Keyword clarity
7. **BUG-004** (Nam - 2h) ← Error handling

---

## ✅ TESTING CHECKLIST

After fixing each bug, test:

### Full Flow Test
```
1. Fresh install
   → pip install -r requirements.txt ✅

2. Run program
   → python main.py ✅

3. "mở powerpoint"
   → PowerPoint opens ✅

4. "tìm file"
   → demo/sample.pptx opens ✅

5. "trình chiếu"
   → Slideshow starts ✅

6. "tiếp tục"
   → Next slide ✅

7. "slide ba"
   → Goes to slide 3 ✅

8. "lùi lại"
   → Previous slide ✅

9. "thoát"
   → Exits cleanly ✅
```

### Edge Cases
```
□ Run without PowerPoint installed
□ Run without microphone
□ Run without internet
□ Say gibberish commands
□ Say commands very fast
□ Say commands during slideshow transition
```

---

## 🆘 IF YOU GET STUCK

### Nam's Bugs (COM issues)
- Check PowerPoint is installed
- Try `win32com.client.gencache.EnsureDispatch()` instead
- Check Windows COM security settings

### Việt's Bugs (Parser issues)
- Print debug info: `print(f"Text: {text}, Matched: {action}")`
- Test each keyword individually
- Check order of commands in config

### Hưng's Bugs (File issues)
- Use absolute paths everywhere
- Print `self.script_dir` to verify
- Check file exists: `if file.exists()`

### Phát's Bugs (Install issues)
- Use fresh venv for testing
- Check Python version: `python --version`
- Try `pip install --upgrade pip` first

---

**Remember**: Test after EACH fix, don't wait until the end!
