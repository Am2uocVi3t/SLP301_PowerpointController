# 👥 TEAM TASK ASSIGNMENTS - Quick Reference

## 🚨 CRITICAL: Read Full Report First!
**Main Report**: `CODE_REVIEW_REPORT.md` (detailed analysis)  
**This Document**: Quick task reference for each team member

---

## 👨‍💻 NAM - Technical Lead (8 hours)

### Your Mission: Fix Critical Controller Bugs

**Skills**: Deep technical, debugging expert  
**Risk**: 🔴 HIGH - Core features depend on you

### P0 Tasks (Day 1) - 6 hours
1. **Fix _next_slide() bug** (2h)
   - File: `controller/powerpoint_controller.py:159-170`
   - Problem: Wrong SlideShowWindow check
   - Impact: "tiếp tục" command fails

2. **Fix _previous_slide() bug** (1h)
   - File: `controller/powerpoint_controller.py:172-183`
   - Same bug as above

3. **Fix _goto_slide() for slideshow** (2h)
   - File: `controller/powerpoint_controller.py:131-151`
   - Must work in both edit AND slideshow mode

4. **Code review all fixes** (1h)

### P1 Tasks (Day 2) - 2 hours
5. **Improve COM exception handling** (2h)
   - Add better try-except blocks
   - Handle edge cases

### Your Success Metric
✅ All slideshow navigation commands work perfectly

---

## 🔧 VIỆT - System Integration (7 hours)

### Your Mission: Fix Parser & Create Tests

**Skills**: System integration, testing  
**Risk**: 🟡 MEDIUM - Quality improvements

### P1 Tasks (Day 2) - 5 hours
1. **Fix keyword conflict "sau"** (1.5h)
   - File: `config/settings.py:21,38`
   - Remove "sau" from goto_slide keywords
   - Keep only "sáu" (accented)

2. **Improve parser logic** (1.5h)
   - File: `parser/command_parser.py`
   - Better number extraction
   - Handle edge cases

3. **Create test suite** (2h)
   - Create: `tests/test_parser.py`
   - Test all command parsing
   - Test number extraction

### P2 Tasks (Day 3) - 2 hours
4. **Advanced parser improvements** (2h)
   - Command priority
   - Better error messages

### Your Success Metric
✅ No command confusion, tests exist

---

## 📁 HƯNG - Demo Setup (5 hours)

### Your Mission: Make Demo Work

**Skills**: Practical tasks, file management  
**Risk**: 🔴 HIGH - Demo needs files

### P0 Tasks (Day 1) - 3 hours
1. **Create demo PowerPoint file** (1h)
   - Create: `demo/sample.pptx`
   - 7 slides with simple content
   - Title: "Speech PowerPoint Controller Demo"
   - **CRITICAL**: Demo fails without this!

2. **Fix path handling** (2h)
   - File: `controller/powerpoint_controller.py:113-131`
   - Use absolute paths instead of relative
   - Add `self.script_dir` variable

### P1 Tasks (Day 2) - 1 hour
3. **Add noise config** (1h)
   - File: `config/settings.py`
   - Add `'noise_adjustment_duration': 1`

### P2 Tasks (Day 3) - 1 hour
4. **Update documentation** (1h)
   - Update README if needed
   - Document path changes

### Your Success Metric
✅ Demo file exists and "tìm file" command works

---

## ⚙️ PHÁT - Installation Setup (4.5 hours)

### Your Mission: Fix Installation

**Skills**: Setup, configuration, testing  
**Risk**: 🔴 HIGH - Can't install = can't demo

### P0 Tasks (Day 1) - 2.5 hours
1. **Fix requirements.txt** (0.5h)
   - File: `requirements.txt`
   - Remove lines 13-14 (standard-aifc, standard-distutils)
   - Change pywin32==307 to pywin32==306
   - **CRITICAL**: Installation fails without this!

2. **Test fresh installation** (1h)
   - Create new venv
   - pip install -r requirements.txt
   - Verify all imports work

3. **Integration testing support** (1h)
   - Help team test full flow
   - Document issues found

### P1 Tasks (Day 2) - 1 hour
4. **Create install test script** (1h)
   - Create: `tests/test_installation.py`
   - Test all imports
   - Clear pass/fail messages

### P2 Tasks (Day 3) - 1 hour
5. **Create install guide** (1h)
   - Create: `INSTALL_GUIDE.md`
   - Step-by-step setup
   - Troubleshooting section

### Your Success Metric
✅ Fresh install works on clean machine

---

## 📅 DAILY SCHEDULE

### Day 1 - Critical Fixes (Everyone) - 8 hours

**09:00 - 12:00 Morning Sprint**
```
Nam:  Fix next/previous slide bugs
Việt: Review code, plan tests  
Hưng: Create demo file + start path fix
Phát: Fix requirements.txt
```

**12:00 - 14:00 Lunch + Continue**
```
Nam:  Fix goto_slide slideshow mode
Hưng: Finish path fix
Phát: Test installation
```

**14:00 - 17:00 Integration**
```
ALL:  Test together
      Fix issues found
      Verify all 7 commands work
```

**End of Day 1 Goal**: ✅ Can demo basic functionality

---

### Day 2 - Quality (Việt, Nam) - 6 hours

**09:00 - 12:00**
```
Nam:  COM exception handling
Việt: Fix keyword conflicts + parser
Hưng: Add config improvements
Phát: Create install test
```

**14:00 - 18:00**
```
Việt: Create test suite
ALL:  Integration testing
ALL:  Practice demo run
```

**End of Day 2 Goal**: ✅ Confident demo, tests exist

---

### Day 3 - Polish (All) - 4 hours

**09:00 - 12:00**
```
Each: Polish assigned areas
```

**14:00 - 18:00**
```
ALL:  Final rehearsal
ALL:  Prepare Q&A
ALL:  Backup plans ready
```

**End of Day 3 Goal**: ✅ Professional demo ready

---

## 🎯 INDIVIDUAL SUCCESS CHECKLIST

### ✅ NAM Checklist
- [ ] _next_slide() works in slideshow
- [ ] _previous_slide() works in slideshow
- [ ] _goto_slide() works in both modes
- [ ] COM exceptions handled properly
- [ ] Code reviewed by you

### ✅ VIỆT Checklist
- [ ] No "sau" in goto_slide keywords
- [ ] Parser handles edge cases
- [ ] test_parser.py created
- [ ] All tests pass

### ✅ HƯNG Checklist
- [ ] demo/sample.pptx exists (7 slides)
- [ ] "tìm file" command finds it
- [ ] Paths use absolute, not relative
- [ ] Config has noise_adjustment_duration

### ✅ PHÁT Checklist
- [ ] requirements.txt has only valid packages
- [ ] Fresh install works
- [ ] test_installation.py created
- [ ] INSTALL_GUIDE.md written

---

## 🐛 CRITICAL BUGS TO FIX

| Bug | Who | Priority | Hours |
|-----|-----|----------|-------|
| BUG-001: No demo file | Hưng | 🔴 P0 | 1h |
| BUG-002: SlideShowWindow check | Nam | 🔴 P0 | 3h |
| BUG-005: requirements.txt | Phát | 🔴 P0 | 0.5h |
| BUG-003: Keyword conflict | Việt | 🟡 P1 | 1.5h |
| BUG-004: COM exceptions | Nam | 🟡 P1 | 2h |
| BUG-006: Path handling | Hưng | 🟡 P1 | 2h |

---

## 💬 COMMUNICATION PROTOCOL

**Daily Standup** (10 mins each morning):
- What I did yesterday
- What I'll do today
- Any blockers

**When Blocked**:
1. Try for 30 minutes
2. Ask in team chat
3. If urgent, call Nam

**When Done with Task**:
1. Test it yourself
2. Commit to git
3. Notify team
4. Help others

**Git Workflow**:
```bash
# Create your branch
git checkout -b nam/fix-slideshow
git checkout -b viet/fix-parser
git checkout -b hung/demo-file
git checkout -b phat/fix-requirements

# Make changes, then:
git add .
git commit -m "Fix: description"
git push origin your-branch-name
```

---

## 🆘 EMERGENCY CONTACTS

**Technical Issues**: Nam (team lead)  
**Integration Issues**: Việt  
**File/Setup Issues**: Hưng  
**Installation Issues**: Phát

---

## 🎓 FINAL TIPS

**For Nam**:
- Your bugs are hardest - start early
- Don't hesitate to refactor if needed
- Test in both edit AND slideshow mode

**For Việt**:
- Write tests as you fix bugs
- Think about edge cases
- Document test scenarios

**For Hưng**:
- Demo file is CRITICAL - do this first
- Make slides clear and simple
- Test file opening multiple times

**For Phát**:
- Test on fresh environment
- Document EVERY step
- Keep installation simple

---

## 📊 PROGRESS TRACKING

Update daily:
```
Day 1: Nam [██░░] 50%  Việt [█░░░] 25%  Hưng [███░] 75%  Phát [████] 100%
```

---

**REMEMBER**: 
- Day 1 = Must work
- Day 2 = Should work well  
- Day 3 = Professional polish

**Good luck team! 🚀**
