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
    return segs, faces_now(ns)


def view_faces(ns, yaw=-35, pitch=25):
    """VBA ViewProject natijasi: [(ekran nuqtalari, yorug'lik)] chizish tartibida"""
    m = ns["G_mod3d"]
    if m.fn == 0:
        return []
    ns["mod3d__viewproject"](yaw, pitch)
    out = []
    for k in range(m.fn):
        f = m.ford(k)
        s, n = m.fstart(f), m.fcnt(f)
        out.append(([(m.vx(s + j), m.vy(s + j)) for j in range(n)], m.flit(f)))
    return out


def faces_now(ns):
    m = ns["G_mod3d"]
    faces = []
    for f in range(m.fn):
        s, n = m.fstart(f), m.fcnt(f)
        faces.append([(m.px(s + k), m.py(s + k), m.pz(s + k)) for k in range(n)])
    return faces


def fold_states(ns):
    """Buklanadigan shablon uchun: [(nom, yuzalar)] - 50%, 100%, qopqoq ochiq"""
    if not ns["G_mod3d"].kinematic:
        return []
    out = []
    for name, t, o in (("50%", 0.5, 0), ("100%", 1, 0), ("ochiq", 1, 1)):
        ns["mod3d__fold3d"](t, o)
        out.append((name, view_faces(ns)))
    ns["mod3d__fold3d"](1, 0)
    return out


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


def iso(g, faces, ox, oy, size, label=None):
    """VBA ViewProject bergan yuzalarni (ox, oy) dagi size x size katakka chizish"""
    if not faces:
        return
    xs = [p[0] for f, _ in faces for p in f]
    ys = [p[1] for f, _ in faces for p in f]
    kk = (size - 40) / max(max(xs) - min(xs), max(ys) - min(ys), 1)
    for f, lit in faces:
        pts = [(ox + 20 + (x - min(xs)) * kk, oy + size - 20 - (y - min(ys)) * kk) for x, y in f]
        if len(pts) >= 3:
            g.polygon(pts, fill=(int(207 * lit), int(169 * lit), int(126 * lit)), outline=(90, 64, 38))
    if label:
        g.text((ox + 5, oy + 5), label, fill="black")


def iso_old(g, faces, ox, oy, size, label=None):
    if not faces:
        return
    cy, sy = math.cos(math.radians(-35)), math.sin(math.radians(-35))
    cp, sp = math.cos(math.radians(25)), math.sin(math.radians(25))

    def rot(v):
        x, y, z = v
        ax, az = x * cy + z * sy, -x * sy + z * cy
        return ax, y * cp - az * sp, y * sp + az * cp
    rf = [[rot(v) for v in f] for f in faces]
    ax = [v[0] for f in rf for v in f]
    ay = [v[1] for f in rf for v in f]
    kk = (size - 40) / max(max(ax) - min(ax), max(ay) - min(ay), 1)
    for f in sorted(rf, key=lambda f: sum(v[2] for v in f) / len(f)):
        pts = [(ox + 20 + (v[0] - min(ax)) * kk, oy + size - 20 - (v[1] - min(ay)) * kk) for v in f]
        (a0, b0, c0), (a1, b1, c1), (a2, b2, c2) = f[0], f[1], f[2]
        nz = (a1 - a0) * (b2 - b0) - (b1 - b0) * (a2 - a0)
        col = (207, 169, 126) if nz >= 0 else (176, 140, 100)
        g.polygon(pts, fill=col, outline=(90, 64, 38))
    if label:
        g.text((ox + 5, oy + 5), label, fill="black")


def tile(segs, faces, title, size=500, states=()):
    from PIL import Image, ImageDraw
    cells = 2 + max(0, len(states) - 1)
    im = Image.new("RGB", (size * cells, size), "white")
    g = ImageDraw.Draw(im)
    xs = [v for s in segs for v in (s[1], s[3])]
    ys = [v for s in segs for v in (s[2], s[4])]
    x0, y1 = min(xs), max(ys)
    k = (size - 30) / max(max(xs) - x0, y1 - min(ys), 1)
    for kind, a, b, c, e in segs:
        P = lambda x, y: (15 + (x - x0) * k, 25 + (y1 - y) * k)
        g.line([P(a, b), P(c, e)], fill="#005AFF" if kind == "CUT" else "#E60028", width=2 if kind == "CUT" else 1)
    if states:
        for j, (name, fs) in enumerate(states):
            iso(g, fs, size * (j + 1) if j < 3 else size, 0, size, name)
    else:
        iso(g, faces, size, 0, size, "sodda 3D")  # faces = view_faces(...)
    g.text((5, 5), title, fill="black")
    return im


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "preview"
    os.makedirs(out, exist_ok=True)
    import glob
    vba2py.lint(sorted(glob.glob(os.path.join(ROOT, "*.bas"))) + glob.glob(os.path.join(ROOT, "*.cls")) +
                glob.glob(os.path.join(ROOT, "*.frm")))   # forma va Corel kodi ham tekshiriladi
    print("lint: xato yo'q")
    ns = load()
    n = ns["modtemplates__tcount"]()
    tiles, bad = [], 0
    for i in range(1, n + 1):
        title = "%d. %s / %s" % (i, ns["modtemplates__tgroup"](i), ns["modtemplates__tname"](i))
        try:
            segs, _ = build(ns, i)
            faces = view_faces(ns)
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
            tiles.append(tile(segs, faces, title, 400, fold_states(ns)))
        except ImportError:
            pass
    if tiles:
        from PIL import Image
        per = 5
        for s in range(0, len(tiles), per):
            chunk = tiles[s:s + per]
            sheet = Image.new("RGB", (1600, 400 * len(chunk)), "#DDDDDD")
            for j, t in enumerate(chunk):
                sheet.paste(t, (0, j * 400))
            sheet.save(os.path.join(out, "sheet-%02d.png" % (s // per + 1)))
    print("%d shablon, %d xato" % (n, bad))
    sys.exit(1 if bad else 0)
