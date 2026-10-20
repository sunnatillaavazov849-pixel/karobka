"""VBA shablonlarining Python nusxasi: geometriyani Corelsiz tekshirish uchun.

Ishlatish: python3 tools/preview.py [papka]   ->  papka/<shablon>.svg (standart: preview)
src/modTemplates.bas o'zgarsa, shu fayl ham bir xil o'zgartiriladi.
"""
import math
import os
import sys

E = 1e-6


def r(v):
    return round(v, 3)


class Die:
    def __init__(self):
        self.edges, self.slits, self.free = [], [], []

    def poly(self, pts):
        n = len(pts)
        for i in range(n):
            a, b = pts[i], pts[(i + 1) % n]
            self.edges.append((r(a[0]), r(a[1]), r(b[0]), r(b[1])))

    def rect(self, x1, y1, x2, y2):
        self.poly([(x1, y1), (x2, y1), (x2, y2), (x1, y2)])

    def slit(self, x1, y1, x2, y2):
        self.slits.append((r(x1), r(y1), r(x2), r(y2)))

    def cut_poly(self, pts):
        n = len(pts)
        for i in range(n):
            a, b = pts[i], pts[(i + 1) % n]
            self.free.append(("CUT", r(a[0]), r(a[1]), r(b[0]), r(b[1])))

    def crease(self, x1, y1, x2, y2):
        self.free.append(("FOLD", r(x1), r(y1), r(x2), r(y2)))

    def solve(self):
        pts = set()
        for e in self.edges + self.slits:          # tirqish uchlari ham bo'luvchi nuqta
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
                if E < t < 1 - E and abs(x1 + t * dx - px) < 1e-3 and abs(y1 + t * dy - py) < 1e-3:
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
        return out + self.free


def on_seg(px, py, s):
    x1, y1, x2, y2 = s
    dx, dy = x2 - x1, y2 - y1
    ln2 = dx * dx + dy * dy
    if ln2 < E:
        return False
    t = ((px - x1) * dx + (py - y1) * dy) / ln2
    return -E <= t <= 1 + E and abs(x1 + t * dx - px) < 1e-3 and abs(y1 + t * dy - py) < 1e-3


def arc(pts, cx, cy, rad, a0, a1, n=6):
    """Yoy nuqtalari (gradus), uchlari bilan"""
    for i in range(n + 1):
        a = math.radians(a0 + (a1 - a0) * i / n)
        pts.append((cx + rad * math.cos(a), cy + rad * math.sin(a)))


# ---------------------------------------------------------------- yordamchilar

def glue(d, x, H, g, s=-1):
    """Yelim qanoti: x chizig'iga ulangan, ~15 gradus qiya"""
    c = min(g * 0.3, H / 4)
    d.poly([(x, 0), (x + s * g, c), (x + s * g, H - c), (x, H)])


def body(d, H, g, widths):
    xs = [0]
    for w in widths:
        xs.append(xs[-1] + w)
    glue(d, 0, H, g)
    for i in range(len(widths)):
        d.rect(xs[i], 0, xs[i + 1], H)
    return xs


def slot_flaps(d, xs, y, s, f):
    for i in range(len(xs) - 1):
        d.rect(xs[i], y, xs[i + 1], y + s * f)
    for x in xs[1:-1]:
        d.slit(x, y, x, y + s * f)


def tuck_closure(d, xs, b, s, lid, L, W, tuck):
    """lid - qopqoq paneli. Quloqlarning qopqoqqa yaqin tomoni tik, uzoq tomoni qiya"""
    dh = min(W * 0.62, L / 2 - 1)
    for i in range(4):
        xa, xb = xs[i], xs[i + 1]
        if i == lid:
            y1, y2 = b + s * W, b + s * (W + tuck)
            rr = min(tuck * 0.7, (xb - xa) / 4)
            d.rect(xa, b, xb, y1)
            pts = [(xa, y1)]
            if s > 0:
                arc(pts, xa + rr, y2 - rr, rr, 180, 90)
                arc(pts, xb - rr, y2 - rr, rr, 90, 0)
            else:
                arc(pts, xa + rr, y2 + rr, rr, 180, 270)
                arc(pts, xb - rr, y2 + rr, rr, 270, 360)
            pts.append((xb, y1))
            d.poly(pts)
        elif i % 2 == 1:
            near_left = (i - 1) % 4 == lid          # qopqoq chap qo'shnimi
            w = xb - xa
            c, step, rr, gap = w * 0.2, dh * 0.12, min(dh * 0.15, w * 0.1), 0.5
            # chapdan o'ngga: tik tomon (gap bilan) yoki qiya tomon
            if near_left:
                left = [(xa, b), (xa + gap, b + s * (dh - rr))]
                right = [(xb - c + rr, b + s * dh), (xb, b + s * step), (xb, b)]
                pts = left
                arc(pts, xa + gap + rr, b + s * (dh - rr), rr, 180, 90 if s > 0 else 270)
                pts += right
            else:
                pts = [(xa, b), (xa, b + s * step), (xa + c - rr, b + s * dh)]
                arc(pts, xb - gap - rr, b + s * (dh - rr), rr, 90 if s > 0 else 270, 0 if s > 0 else 360)
                pts += [(xb - gap, b + s * (dh - rr)), (xb, b)]
            d.poly(pts)


def lock_wall(x0, sx, W, H, t):
    """Ikki qavat yon devorning ichki qismi (qulf tillari bilan) nuqtalari. x0 - devor fold chizig'i"""
    xi, xe = x0 + sx * H, x0 + sx * (2 * H - t)
    td = max(2 * t, 3)
    pts = [(xi, 0), (xe, 0)]
    for a, b in ((0.2, 0.4), (0.6, 0.8)):
        pts += [(xe, a * W), (xe + sx * td, a * W + 1), (xe + sx * td, b * W - 1), (xe, b * W)]
    pts += [(xe, W), (xi, W)]
    return pts


def base_slots(d, x, W):
    for a, b in ((0.2, 0.4), (0.6, 0.8)):
        d.slit(x, a * W, x, b * W)


# ---------------------------------------------------------------- shablonlar

def rsc(p, top=True):
    d = Die()
    xs = body(d, p["H"], p["G"], [p["L"], p["W"], p["L"], p["W"]])
    if top:
        slot_flaps(d, xs, p["H"], 1, p["W"] / 2)
    slot_flaps(d, xs, 0, -1, p["W"] / 2)
    return d


def tuck(p, bottom_lid):
    d = Die()
    L, W, H = p["L"], p["W"], p["H"]
    xs = body(d, H, p["G"], [L, W, L, W])
    tuck_closure(d, xs, H, 1, 0, L, W, p["TK"])
    tuck_closure(d, xs, 0, -1, bottom_lid, L, W, p["TK"])
    return d


def mailer(p, window=False):
    d = Die()
    L, W, H, t = p["L"], p["W"], p["H"], p["T"]
    e_len = min(1.3 * H, W / 2 - 2)
    d.rect(0, 0, L, W)                                   # taglik
    d.rect(0, -H, L, 0)                                  # old devor
    d.rect(0, W, L, W + H)                               # orqa devor
    for x0, sx in ((0, -1), (L, 1)):
        d.rect(x0, 0, x0 + sx * H, W)                    # yon devor
        d.poly(lock_wall(x0, sx, W, H, t))               # ichki yon devor
        base_slots(d, x0, W)
        for y0, sy in ((0, -1), (W, 1)):                 # quloqlar old/orqa devorga ulangan
            d.rect(x0, y0, x0 + sx * e_len, y0 + sy * H)
            d.slit(x0, y0, x0 + sx * e_len, y0)
    d.rect(0, W + H, L, 2 * W + H)                       # qopqoq
    dd, k = H - 2 * t, (H - 2 * t) * 0.35
    for x0, sx in ((0, -1), (L, 1)):                     # qopqoq quloqlari
        d.poly([(x0, W + H + t), (x0 + sx * dd, W + H + t + k), (x0 + sx * dd, 2 * W + H - t - k), (x0, 2 * W + H - t)])
    y0, ht = 2 * W + H, H - t                            # old til
    d.rect(0, y0, L, y0 + ht)
    e = H - 2 * t
    rr = min(e, ht) * 0.9
    for x0, sx in ((0, -1), (L, 1)):
        pts = [(x0, y0), (x0 + sx * e, y0)]
        if sx < 0:
            arc(pts, x0 - e + rr, y0 + ht - rr, rr, 180, 90)
        else:
            arc(pts, x0 + e - rr, y0 + ht - rr, rr, 0, 90)
        pts.append((x0, y0 + ht))
        d.poly(pts)
    if window:
        m = min(p["M"], min(L, W) / 2 - 5)
        d.cut_poly([(m, W + H + m), (L - m, W + H + m), (L - m, 2 * W + H - m), (m, 2 * W + H - m)])
    return d


def tray_into(d, ox, L, W, H):
    """Oddiy lotok: quloqlar yon devorlarga ulangan, old/orqa devorga yelimlanadi"""
    e_len = min(H, L / 2 - 2)
    d.rect(ox, 0, ox + L, W)
    d.rect(ox, -H, ox + L, 0)
    d.rect(ox, W, ox + L, W + H)
    for x0, sx in ((ox, -1), (ox + L, 1)):
        d.rect(x0, 0, x0 + sx * H, W)
        for y0, sy in ((0, -1), (W, 1)):
            d.poly([(x0, y0), (x0 + sx * (H - 1), y0), (x0 + sx * (H - 1), y0 + sy * (e_len - 2)), (x0 + sx * 2, y0 + sy * e_len), (x0, y0 + sy * e_len)])
            d.slit(x0, y0, x0, y0 + sy * e_len)


def tray(p):
    d = Die()
    tray_into(d, 0, p["L"], p["W"], p["H"])
    return d


def tray_double(p):
    """Ikki qavat old/orqa devor (qulf tillari), quloqlar yon devorlarda"""
    d = Die()
    L, W, H, t = p["L"], p["W"], p["H"], p["T"]
    e_len = min(1.3 * H, L / 2 - 2)
    d.rect(0, 0, L, W)
    for x0, sx in ((0, -1), (L, 1)):
        d.rect(x0, 0, x0 + sx * H, W)
        for y0, sy in ((0, -1), (W, 1)):
            d.rect(x0, y0, x0 + sx * (H - 1), y0 + sy * e_len)
            d.slit(x0, y0, x0, y0 + sy * e_len)
    for y0, sy in ((0, -1), (W, 1)):
        d.rect(0, y0, L, y0 + sy * H)
        pts = [(x, y) for (y, x) in lock_wall(y0, sy, L, H, t)]   # o'qlar almashgan
        d.poly(pts)
        for a, b in ((0.2, 0.4), (0.6, 0.8)):
            d.slit(a * L, y0, b * L, y0)
    return d


def telescope(p):
    d = Die()
    L, W, H, hl, k = p["L"], p["W"], p["H"], p["HL"], p["K"]
    tray_into(d, 0, L, W, H)
    tray_into(d, L + 2 * H + 20 + hl, L + 2 * k, W + 2 * k, hl)
    return d


TEMPLATES = {
    "fefco-0201": (rsc, dict(L=300, W=200, H=150, G=35)),
    "tuck-reverse": (lambda p: tuck(p, 2), dict(L=80, W=50, H=150, G=15, TK=15)),
    "tuck-straight": (lambda p: tuck(p, 0), dict(L=80, W=50, H=150, G=15, TK=15)),
    "mailer": (mailer, dict(L=250, W=180, H=60, T=2)),
    "mailer-window": (lambda p: mailer(p, True), dict(L=250, W=180, H=60, T=2, M=35)),
    "tray": (tray, dict(L=200, W=150, H=50)),
    "tray-double": (tray_double, dict(L=250, W=180, H=50, T=2)),
    "telescope": (telescope, dict(L=200, W=150, H=50, HL=30, K=2)),
}


def svg(segs, path):
    xs = [v for s in segs for v in (s[1], s[3])]
    ys = [v for s in segs for v in (s[2], s[4])]
    x0, y0, x1, y1 = min(xs) - 10, min(ys) - 10, max(xs) + 10, max(ys) + 10
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{x0} {-y1} {x1 - x0} {y1 - y0}" '
           f'width="{(x1 - x0) * 2}" height="{(y1 - y0) * 2}"><rect x="{x0}" y="{-y1}" width="100%" height="100%" fill="white"/>']
    for k, a, b, c, e in segs:
        st = 'stroke="#005AFF"' if k == "CUT" else 'stroke="#E60028" stroke-dasharray="3 2"'
        out.append(f'<line x1="{a}" y1="{-b}" x2="{c}" y2="{-e}" {st} stroke-width="0.6"/>')
    out.append("</svg>")
    with open(path, "w") as f:
        f.write("\n".join(out))


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "preview"
    os.makedirs(out, exist_ok=True)
    for name, (fn, params) in TEMPLATES.items():
        segs = fn(params).solve()
        svg(segs, os.path.join(out, name + ".svg"))
        print(name, len(segs))
