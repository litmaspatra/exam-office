# Antigravity prompt: Exam Office (standalone RSSB typing test app)

Use this if you want Antigravity to regenerate or extend the app instead of using the delivered exam-office.html directly. One pass:

---

Create a standalone single-file web app `exam-office.html` (no build step, no server; opens by double-click). It is a Soni Typing Tutor-style exam center for the NEW RSSB Clerk Grade-II / Junior Assistant Phase-II pattern. Reuse the KrutiDev010 base64 @font-face, the 4-theme CSS variable system (dark/navy/amber/light), the `unicodeToKrutiDev()` converter, the NUMPAD_DIGIT_MAP + KRUTIDEV_ALT_CODES Alt+numpad entry, and the WebAudio click/error sounds from hikeybr's index.html. Do NOT reuse hikeybr's practice/progression/mastery logic - this is a separate app. IMPORTANT: hikeybr's unicodeToKrutiDev lacks छ and झ - add them: छ -> 'N', झ -> '>' (per its own KEY_PROGRESSION).

VERIFIED RSSB PATTERN (implement exactly, no WPM targets, no negative marking):
- 4 papers x 10 min x 25 marks; minimum 9 marks in EACH paper (36%); SC/ST get 5% relaxation (7.75).
- Paper 1 Hindi speed: 400-word passage, DevLys 010 font, Remington layout. Marks = correct words x 0.0625 (max 25). 144 correct words = 9 marks (~14.4 WPM); 400 = 25 marks (40 WPM).
- Paper 2 English speed: 500 words, QWERTY, 0.05/word (max 25). 180 = 9 marks; 500 = 25.
- Papers 3-4 (Hindi/English efficiency) are MS Word 2013 / MS Excel 2013 formatting tasks - NOT simulated; show an info card telling the user to practise them in Word/Excel.
- A "word" = characters between spaces. A word earns 0 if it has a wrong/missing/repeated character, word or digit, punctuation mismatch, spacing error, wrong case (English), or wrong/invisible-duplicate matra (Hindi).
- Exam hall rule to enforce: Backspace works ONLY inside the current word; once Space is pressed the word is final.

FEATURES:
1. Tests home: cards for 3 Hindi tests (~400 words each) and 2 English tests (~500 words each), each showing word count, 10 min, per-word marks, and the user's best marks so far. Header brand "Exam Office", nav tabs: Tests / My Results / Exam Pattern. Dark theme default with theme cycler and sound toggle (persisted).
2. Test runner: passage rendered in the embedded KrutiDev font for Hindi (plain mono for English), one span per word. Timer starts on the FIRST keystroke, 10:00 countdown, auto-submit at 0. Live chips: words done, correct, wrong, live marks, net WPM, accuracy. Per-character coloring inside the current word (correct/wrong/pending + blinking caret); committed words turn green (correct) or red (wrong/skipped). Typing captures physical keys on a hidden input; Hindi keystrokes are raw KrutiDev bytes (event.key with e.code overrides for punctuation; Space commits the word, even empty = skipped; Backspace deletes within the current word only; Alt+numpad digits buffered and converted on Alt release via KRUTIDEV_ALT_CODES + unicodeToKrutiDev; block paste, Tab, Enter, Ctrl/Meta). Compare committed word bytes EXACTLY against unicodeToKrutiDev(expectedWord). End on: time up / Submit button / last word committed. Quit asks confirm and discards.
3. Result screen (Soni-style report): big marks /25, QUALIFIED (>=9) or NOT QUALIFIED badge with "need 9.00 Gen/OBC, 7.75 SC/ST" note (when qualified show "cleared by X marks", when failed show "N more correct words to qualify"), stat cards: net WPM (correct/elapsed min), gross WPM (attempted/elapsed), accuracy, correct/wrong/omitted counts, time used. Word-by-word table of wrong + skipped words (expected in Unicode Devanagari, typed rendered in KrutiDev font; use a plain '-' placeholder for skipped, never in KrutiDev font) + "N words not reached" summary row. Buttons: Retry / Back to Tests. Save every attempt to localStorage (key exam_office_store_v1) and show a My Results history table (date, test, marks, net WPM, accuracy, qualified) with a clear-history button.
4. Exam Pattern tab: the 4-paper table, the marking formula, the 0-marks error list (include the invisible-matra examples: केवल = dsoy correct, dsssssoy wrong; प = i, not Ik), exam-hall rules, and a source line pointing to rssb.rajasthan.gov.in.

PASSAGE RULES: Hindi passages in formal exam-style Unicode Hindi, 400-430 words, using ONLY characters the converter supports (letters, matras, अं अः, danda ।, Devanagari digits; no ASCII punctuation; avoid ऑ which KrutiDev010 lacks). English passages plain ASCII, 500-510 words. Verify with a script that unicodeToKrutiDev(word) yields only ASCII 32-126 or Alt bytes 128-255 for every word before shipping.
