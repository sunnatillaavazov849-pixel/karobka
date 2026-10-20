"""src/*.bas dagi shablonlarni Corelsiz chizib tekshirish.

Ishlatish:  python3 tools/preview.py [papka]
  papka/NN-nom.svg  - har bir shablon yoyilmasi
  papka/sheet-*.png - hammasi bitta varaqda (yoyilma + 3D)
VBA kodi tools/vba2py.py orqali Python'ga o'giriladi, shuning uchun alohida nusxa yozish shart emas.
"""
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
import vba2py  # noqa: E402

ROOT = os.path.join(os.path.dirname(__file__), "..", "src")


def load():
    files = [os.path.join(ROOT, f) for f in ("modGeom.bas", "mod3D.bas", "modTemplates.bas")]
    ns, _ = vba2py.load(files, skip=("Draw3D",))
    return ns


def build(ns, i, overrides=None):
    p = vba2py.VDict()
    for item in str(ns["modtemplates__tparams"](i)).split("|"):
        parts = item.split("=")
        p[parts[0]] = float(parts[2])
    for k, v in (overrides or {}).items():
        p[k] = v
    ns["modtemplates__tbuild"](i, p)
    g = ns["G_modgeom"]
    segs = [("CUT" if g.segkind(j) == 1 else "FOLD", g.segx1(j), g.segy1(j), g.segx2(j), g.segy2(j))
            for j in range(g.segn)]
    m = ns["G_mod3d"]
    faces = []
    for f in range(m.fn):
        s, n = m.fstart(f), m.fcnt(f)
        faces.append([(m.px(s + k), m.py(s + k), m.pz(s + k)) for k in range(n)])
    return segs, faces


def check(segs):
    """Muammolar ro'yxati: nol uzunlik, takroriy chiziq"""
    probs, seen = [], set()
    for k, a, b, c, d in segs:
        if abs(a - c) < 1e-6 and abs(b - d) < 1e-6:
            probs.append("nol uzunlikdagi chiziq")
        key = (min((a, b), (c, d)), max((a, b), (c, d)))
        if key in seen:
            probs.append("takroriy chiziq %s" % (key,))
        seen.add(key)
    return probs


def svg(segs, path):
    xs = [v for s in segs for v in (s[1], s[3])]
    ys = [v for s in segs for v in (s[2], s[4])]
    x0, y0, x1, y1 = min(xs) - 10, min(ys) - 10, max(xs) + 10, max(ys) + 10
    out = ['<svg xmlns="http://www.w3.org/2000/svg" viewBox="%g %g %g %g" width="%g" height="%g">'
           '<rect x="%g" y="%g" width="100%%" height="100%%" fill="white"/>'
           % (x0, -y1, x1 - x0, y1 - y0, (x1 - x0) * 2, (y1 - y0) * 2, x0, -y1)]
    for k, a, b, c, e in segs:
        st = 'stroke="#005AFF"' if k == "CUT" else 'stroke="#E60028" stroke-dasharray="3 2"'
        out.append('<line x1="%g" y1="%g" x2="%g" y2="%g" %s stroke-width="0.6"/>' % (a, -b, c, -e, st))
    out.append("</svg>")
    with open(path, "w") as f:
        f.write("\n".join(out))


def tile(segs, faces, title, size=500):
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (size * 2, size), "white")
    g = ImageDraw.Draw(im)
    xs = [v for s in segs for v in (s[1], s[3])]
    ys = [v for s in segs for v in (s[2], s[4])]
    x0, y1 = min(xs), max(ys)
    k = (size - 30) / max(max(xs) - x0, y1 - min(ys), 1)
    for kind, a, b, c, e in segs:
        P = lambda x, y: (15 + (x - x0) * k, 25 + (y1 - y) * k)
        g.line([P(a, b), P(c, e)], fill="#005AFF" if kind == "CUT" else "#E60028", width=2 if kind == "CUT" else 1)
    if faces:
        cy, sy = math.cos(math.radians(-35)), math.sin(math.radians(-35))
        cp, sp = math.cos(math.radians(25)), math.sin(math.radians(25))

        def rot(v):
            x, y, z = v
            ax, az = x * cy + z * sy, -x * sy + z * cy
            return ax, y * cp - az * sp, y * sp + az * cp
        rf = [[rot(v) for v in f] for f in faces]
        ax = [v[0] for f in rf for v in f]
        ay = [v[1] for f in rf for v in f]
        kk = (size - 60) / max(max(ax) - min(ax), max(ay) - min(ay), 1)
        for f in sorted(rf, key=lambda f: sum(v[2] for v in f) / len(f)):
            pts = [(size + 30 + (v[0] - min(ax)) * kk, size - 30 - (v[1] - min(ay)) * kk) for v in f]
            g.polygon(pts, fill=(207, 169, 126), outline=(90, 64, 38))
    g.text((5, 5), title, fill="black")
    return im


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "preview"
    os.makedirs(out, exist_ok=True)
    ns = load()
    n = ns["modtemplates__tcount"]()
    tiles, bad = [], 0
    for i in range(1, n + 1):
        title = "%d. %s / %s" % (i, ns["modtemplates__tgroup"](i), ns["modtemplates__tname"](i))
        try:
            segs, faces = build(ns, i)
            probs = check(segs)
        except Exception as ex:  # noqa: BLE001
            print("XATO", title, ex)
            bad += 1
            continue
        if probs:
            print("OGOH", title, probs[:3])
        fname = re.sub(r"[^\w]+", "-", "%02d-%s" % (i, ns["modtemplates__tname"](i))).strip("-")
        svg(segs, os.path.join(out, fname + ".svg"))
        try:
            tiles.append(tile(segs, faces, title))
        except ImportError:
            pass
    if tiles:
        from PIL import Image
        per = 8
        for s in range(0, len(tiles), per):
            chunk = tiles[s:s + per]
            sheet = Image.new("RGB", (2000, 500 * ((len(chunk) + 1) // 2)), "#DDDDDD")
            for j, t in enumerate(chunk):
                sheet.paste(t, ((j % 2) * 1000, (j // 2) * 500))
            sheet.save(os.path.join(out, "sheet-%02d.png" % (s // per + 1)))
    print("%d shablon, %d xato" % (n, bad))
    sys.exit(1 if bad else 0)
