"""src/* dan nusxalab joylash uchun toza kod (src/joylash/*.txt) yasaydi.

VBA kod oynasiga VERSION / Begin / Attribute qatorlari tushsa
"Invalid outside procedure" xatosi chiqadi, shuning uchun ular olib tashlanadi.
Ishlatish: python3 tools/make_paste.py
"""
import glob
import os

os.makedirs("src/joylash", exist_ok=True)
for path in sorted(glob.glob("src/*.bas") + glob.glob("src/*.cls") + glob.glob("src/*.frm")):
    lines = open(path, newline="").read().replace("\r\n", "\n").split("\n")
    start = next(i for i, ln in enumerate(lines) if not ln.startswith(("VERSION", "Begin", "End", "Attribute", "  ", "   ", "BEGIN", "END")))
    body = "\r\n".join(lines[start:])
    name = os.path.splitext(os.path.basename(path))[0]
    open(f"src/joylash/{name}.txt", "w", newline="").write(body)
    print(name, "->", start, "qator olib tashlandi")
