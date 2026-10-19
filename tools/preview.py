"""Karobka shablonlarini SVG ga chizish (VBA kodi bilan bir xil mantiq, tekshirish uchun).

Ishlatish: python3 tools/preview.py  ->  preview/*.svg
"""
import os

E = 1e-6


def r(v):
    return round(v, 3)


class Die:
    """Panellar (ko'pburchaklar) va majburiy kesish chiziqlari.

    Ikki panelga umumiy qirra -> bukish (FOLD), bitta panelga tegishli -> kesish (CUT).
    Slit (majburiy kesish) bilan ustma-ust tushgan qirra -> CUT.
    """

    def __init__(self):
        self.edges = []
        self.slits = []

    def poly(self, *pts):
        n = len(pts)
        for i in range(n):
            a, b = pts[i], pts[(i + 1) % n]
            self.edges.append((r(a[0]), r(a[1]), r(b[0]), r(b[1])))

    def rect(self, x1, y1, x2, y2):
        self.poly((x1, y1), (x2, y1), (x2, y2), (x1, y2))

    def slit(self, x1, y1, x2, y2):
        self.slits.append((r(x1), r(y1), r(x2), r(y2)))

    def solve(self):
        pts = set()
        for e in self.edges:
            pts.add((e[0], e[1]))
            pts.add((e[2], e[3]))
        pieces = {}
        for x1, y1, x2, y2 in self.edges:
            dx, dy = x2 - x1, y2 - y1
            ln2 = dx * dx + dy * dy
            if ln2 < E:
                continue
            ts = [0.0, 1.0]
            for px, py in pts:
                t = ((px - x1) * dx + (py - y1) * dy) / ln2
                if E < t < 1 - E:
                    cx, cy = x1 + t * dx, y1 + t * dy
                    if abs(cx - px) < 1e-3 and abs(cy - py) < 1e-3:
                        ts.append(t)
            ts.sort()
            for a, b in zip(ts, ts[1:]):
                if b - a < E:
                    continue
                p = (r(x1 + a * dx), r(y1 + a * dy))
                q = (r(x1 + b * dx), r(y1 + b * dy))
                k = (p, q) if p <= q else (q, p)
                pieces[k] = pieces.get(k, 0) + 1
        out = []
        for (p, q), cnt in pieces.items():
            kind = "FOLD" if cnt >= 2 else "CUT"
            mx, my = (p[0] + q[0]) / 2, (p[1] + q[1]) / 2
            if kind == "FOLD" and any(on_seg(mx, my, s) for s in self.slits):
                kind = "CUT"
            out.append((kind, p[0], p[1], q[0], q[1]))
        return out


def on_seg(px, py, s):
    x1, y1, x2, y2 = s
    dx, dy = x2 - x1, y2 - y1
    ln2 = dx * dx + dy * dy
    t = ((px - x1) * dx + (py - y1) * dy) / ln2
    if t < -E or t > 1 + E:
        return False
    return abs(x1 + t * dx - px) < 1e-3 and abs(y1 + t * dy - py) < 1e-3


# ---------------- shablonlar ----------------

def body(d, L, W, H, g):
    """Glue | L | W | L | W yon devorlari, qaytaradi x lar."""
    xs = [0, L, L + W, 2 * L + W, 2 * L + 2 * W]
    c = min(g * 0.6, H / 4)
    d.poly((0, 0), (-g, c), (-g, H - c), (0, H))
    for i in range(4):
        d.rect(xs[i], 0, xs[i + 1], H)
    return xs


def rsc_0201(p):
    L, W, H, g = p["L"], p["W"], p["H"], p["G"]
    d = Die()
    xs = body(d, L, W, H, g)
    f = W / 2
    for i in range(4):
        d.rect(xs[i], H, xs[i + 1], H + f)
        d.rect(xs[i], -f, xs[i + 1], 0)
    for x in xs[1:4]:
        d.slit(x, H, x, H + f)
        d.slit(x, -f, x, 0)
    return d


def tuck_closure(d, xs, base, s, lid_idx, W, tuck, L):
    dh = min(W * 0.8, L / 2)
    c = min(dh / 3, (xs[2] - xs[1]) / 4)
    for i in range(4):
        xa, xb = xs[i], xs[i + 1]
        if i == lid_idx:
            y1, y2 = base + s * W, base + s * (W + tuck)
            tc = min(tuck * 0.8, (xb - xa) / 4)
            d.rect(xa, base, xb, y1)
            d.poly((xa, y1), (xa + tc, y2), (xb - tc, y2), (xb, y1))
        elif i % 2 == 1:
            d.poly((xa, base), (xa + c, base + s * dh), (xb - c, base + s * dh), (xb, base))


def reverse_tuck(p):
    L, W, H, g, t = p["L"], p["W"], p["H"], p["G"], p["T"]
    d = Die()
    xs = body(d, L, W, H, g)
    tuck_closure(d, xs, H, 1, 2, W, t, L)
    tuck_closure(d, xs, 0, -1, 0, W, t, L)
    return d


def straight_tuck(p):
    L, W, H, g, t = p["L"], p["W"], p["H"], p["G"], p["T"]
    d = Die()
    xs = body(d, L, W, H, g)
    tuck_closure(d, xs, H, 1, 2, W, t, L)
    tuck_closure(d, xs, 0, -1, 2, W, t, L)
    return d


def tray(p):
    L, W, H = p["L"], p["W"], p["H"]
    d = Die()
    d.rect(0, 0, L, W)
    d.rect(0, -H, L, 0)
    d.rect(0, W, L, W + H)
    d.rect(-H, 0, 0, W)
    d.rect(L, 0, L + H, W)
    # burchak yelim qanotlari old/orqa devorga ulangan
    c = H * 0.25
    for (x0, sx) in ((0, -1), (L, 1)):
        for (y0, sy) in ((0, -1), (W, 1)):
            d.poly((x0, y0), (x0, y0 + sy * H), (x0 + sx * (H - 2), y0 + sy * (H - c)),
                   (x0 + sx * (H - 2), y0 + sy * c))
            d.slit(x0, y0, x0 + sx * H, y0)
    return d


def tray_lid(p):
    """Telescope: taglik + qopqoq yonma-yon."""
    L, W, H, HL, k = p["L"], p["W"], p["H"], p["HL"], p["K"]
    a = tray({"L": L, "W": W, "H": H})
    b = tray({"L": L + 2 * k, "W": W + 2 * k, "H": HL})
    off = L + 2 * H + 20 + HL
    a.edges += [(e[0] + off, e[1], e[2] + off, e[3]) for e in b.edges]
    a.slits += [(s[0] + off, s[1], s[2] + off, s[3]) for s in b.slits]
    return a


def mailer(p):
    """Mailer (FEFCO 0427 ga o'xshash): qopqoqli, ikki qavat old/yon devor."""
    L, W, H, t = p["L"], p["W"], p["H"], p["T"]
    d = Die()
    d.rect(0, 0, L, W)                     # taglik
    d.rect(0, -H, L, 0)                    # old devor
    d.rect(0, -2 * H + t, L, -H)           # old ichki devor
    d.rect(-H, 0, 0, W)                    # chap devor
    d.rect(-2 * H + t, 0, -H, W)           # chap ichki
    d.rect(L, 0, L + H, W)                 # o'ng devor
    d.rect(L + H, 0, L + 2 * H - t, W)     # o'ng ichki
    d.rect(0, W, L, W + H)                 # orqa devor
    d.rect(0, W + H, L, 2 * W + H)         # qopqoq
    tc = min(H * 0.4, L / 6)
    d.poly((0, 2 * W + H), (tc, 2 * W + 2 * H - t),
           (L - tc, 2 * W + 2 * H - t), (L, 2 * W + H))   # old til
    # qopqoq yon qanotlari
    fc = min(H * 0.5, W / 4)
    for x0, sx in ((0, -1), (L, 1)):
        d.poly((x0, W + H + t), (x0 + sx * (H - t), W + H + t + fc),
               (x0 + sx * (H - t), 2 * W + H - fc), (x0, 2 * W + H - t))
        # quloqlar (old va orqa devorga ulangan)
        for y0, sy in ((0, -1), (W, 1)):
            d.poly((x0, y0 + sy * 0.5), (x0 + sx * (H - 2 * t), y0 + sy * (0.5 + fc * 0.6)),
                   (x0 + sx * (H - 2 * t), y0 + sy * (H - 1)), (x0, y0 + sy * H))
    return d


TEMPLATES = {
    "rsc_0201": (rsc_0201, dict(L=300, W=200, H=150, G=35)),
    "reverse_tuck": (reverse_tuck, dict(L=80, W=50, H=150, G=15, T=15)),
    "straight_tuck": (straight_tuck, dict(L=80, W=50, H=150, G=15, T=15)),
    "tray": (tray, dict(L=200, W=150, H=50)),
    "tray_lid": (tray_lid, dict(L=200, W=150, H=50, HL=30, K=2)),
    "mailer": (mailer, dict(L=250, W=180, H=60, T=2)),
}


def svg(segs, path):
    xs = [v for s in segs for v in (s[1], s[3])]
    ys = [v for s in segs for v in (s[2], s[4])]
    x0, y0, x1, y1 = min(xs) - 10, min(ys) - 10, max(xs) + 10, max(ys) + 10
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{x0} {-y1} {x1 - x0} {y1 - y0}" '
           f'width="{(x1 - x0) * 2}" height="{(y1 - y0) * 2}"><rect x="{x0}" y="{-y1}" width="100%" height="100%" fill="white"/>']
    for k, a, b, c, e in segs:
        st = 'stroke="blue"' if k == "CUT" else 'stroke="red" stroke-dasharray="3 2"'
        out.append(f'<line x1="{a}" y1="{-b}" x2="{c}" y2="{-e}" {st} stroke-width="0.6"/>')
    out.append("</svg>")
    open(path, "w").write("\n".join(out))


if __name__ == "__main__":
    os.makedirs("preview", exist_ok=True)
    for name, (fn, params) in TEMPLATES.items():
        segs = fn(params).solve()
        svg(segs, f"preview/{name}.svg")
        print(name, len(segs), "chiziq")
