"""Taqqoslash varaqlari: chapda asl rasmdagi chizma, o'rtada plagin yoyilmasi, o'ngda 3D.

Ishlatish:  python3 tools/compare.py RASMLAR_PAPKASI CHIQISH_PAPKASI [0311 0340]
  RASMLAR_PAPKASI - photo.rar ochilgan joy (IMG_0311.jpeg ...)
  oxirgi ikki son - rasm raqamlari oralig'i (ixtiyoriy)
Nomi "0311-2 ..." bilan boshlangan shablon IMG_0311 ning 2-qatoriga solishtiriladi.
Rasmda bitta chizma bo'lsa butun rasm olinadi.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
import preview  # noqa: E402
from PIL import Image, ImageDraw  # noqa: E402

CELL = 360


def crop(imgdir, num, row, rows):
    im = Image.open(os.path.join(imgdir, "IMG_%s.jpeg" % num))
    w, h = im.size
    if rows > 1:
        im = im.crop((int(w * 0.3), int(h * (row - 1) / 3), w, int(h * row / 3)))
    im.thumbnail((CELL, CELL))
    return im


def main():
    imgdir, out = sys.argv[1], sys.argv[2]
    lo, hi = (sys.argv[3], sys.argv[4]) if len(sys.argv) > 4 else ("0000", "9999")
    os.makedirs(out, exist_ok=True)
    ns = preview.load()
    n = ns["modtemplates__tcount"]()
    items = []
    for i in range(1, n + 1):
        m = re.match(r"(\d{4})-(\d) ", str(ns["modtemplates__tname"](i)))
        if m and lo <= m.group(1) <= hi:
            items.append((i, m.group(1), int(m.group(2))))
    rows = {}
    for _, num, r in items:
        rows[num] = max(rows.get(num, 0), r)
    tiles = []
    for i, num, r in items:
        name = str(ns["modtemplates__tname"](i))
        segs, _ = preview.build(ns, i)
        states = preview.fold_states(ns)
        if states:
            t = preview.tile(segs, None, name, CELL, [states[1]])
        else:
            t = preview.tile(segs, preview.view_faces(ns), name, CELL)
        cell = Image.new("RGB", (CELL * 3, CELL), "white")
        cell.paste(crop(imgdir, num, r, rows[num]), (0, 0))
        cell.paste(t, (CELL, 0))
        ImageDraw.Draw(cell).text((4, 4), num + "-%d (asl)" % r, fill="red")
        tiles.append(cell)
    per = 6
    for s in range(0, len(tiles), per):
        chunk = tiles[s:s + per]
        sheet = Image.new("RGB", (CELL * 3, CELL * len(chunk) + 2 * len(chunk)), "#BBBBBB")
        for j, t in enumerate(chunk):
            sheet.paste(t, (0, j * (CELL + 2)))
        sheet.save(os.path.join(out, "taqqos-%02d.png" % (s // per + 1)))
    print(len(tiles), "chizma,", (len(tiles) + per - 1) // per, "varaq")


if __name__ == "__main__":
    main()
