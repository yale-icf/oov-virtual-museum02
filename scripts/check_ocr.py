import sqlite3, os
db = sqlite3.connect(os.path.join(os.path.dirname(__file__), "corpus.db"))
total = db.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
with_ocr = db.execute("SELECT COUNT(*) FROM docs WHERE ocr_text != ''").fetchone()[0]
no_ocr = db.execute("SELECT id FROM docs WHERE ocr_text = ''").fetchall()
print(f"Total docs:    {total}")
print(f"With OCR text: {with_ocr}")
print(f"Without OCR:   {len(no_ocr)}")
for row in no_ocr:
    print(f"  - {row[0]}")
