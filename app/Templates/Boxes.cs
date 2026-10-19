using System.Collections.Generic;
using Karobka.Geometry;

namespace Karobka.Templates
{
    public static class Catalog
    {
        public static readonly Template[] All =
        {
            new Rsc0201(), new Hsc0200(), new TuckReverse(), new TuckStraight(), new Gable(), new Pillow(),
            new Mailer(), new MailerWindow(), new Tray(), new TrayDouble(), new Telescope(), new Sleeve(), new Drawer(),
        };
    }

    // ================================================================ Gofra

    public sealed class Rsc0201 : Template
    {
        public override string Id => "fefco-0201";
        public override string Name => "FEFCO 0201";
        public override string Group => "Gofra quti";
        public override Param[] Params => new[] { L(300), W(200), H(150), G(35) };

        public override void Build(Die d, P p)
        {
            var xs = Body(d, p["H"], p["G"], p["L"], p["W"], p["L"], p["W"]);
            SlotFlaps(d, xs, p["H"], 1, p["W"] / 2);
            SlotFlaps(d, xs, 0, -1, p["W"] / 2);
        }

        public override Model3 Model(P p) { var m = new Model3(); m.Box(0, 0, 0, p["L"], p["H"], p["W"]); return m; }
    }

    public sealed class Hsc0200 : Template
    {
        public override string Id => "fefco-0200";
        public override string Name => "FEFCO 0200 (ochiq)";
        public override string Group => "Gofra quti";
        public override Param[] Params => new[] { L(300), W(200), H(150), G(35) };

        public override void Build(Die d, P p)
        {
            var xs = Body(d, p["H"], p["G"], p["L"], p["W"], p["L"], p["W"]);
            SlotFlaps(d, xs, 0, -1, p["W"] / 2);
        }

        public override Model3 Model(P p) { var m = new Model3(); m.Box(0, 0, 0, p["L"], p["H"], p["W"], top: false); return m; }
    }

    // ================================================================ Karton quti (tuck-end)

    public abstract class TuckBase : Template
    {
        public override string Group => "Tuck-end";
        public override Param[] Params => new[] { L(80), W(50), H(150), G(15), Tuck(15) };
        protected abstract int BottomLid { get; }

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], h = p["H"];
            var xs = Body(d, h, p["G"], l, w, l, w);
            Closure(d, xs, h, 1, 2, l, w, p["TK"]);
            Closure(d, xs, 0, -1, BottomLid, l, w, p["TK"]);
        }

        static void Closure(Die d, double[] xs, double b, double s, int lid, double l, double w, double tuck)
        {
            double dh = Min(w * 0.8, l / 2), c = Min(dh / 3, w / 4);
            for (int i = 0; i < 4; i++)
            {
                double xa = xs[i], xb = xs[i + 1];
                if (i == lid)
                {
                    double y1 = b + s * w, y2 = b + s * (w + tuck), tc = Min(tuck * 0.8, (xb - xa) / 4);
                    d.Rect(xa, b, xb, y1);
                    d.Poly(xa, y1, xa + tc, y2, xb - tc, y2, xb, y1);
                }
                else if (i % 2 == 1)
                    d.Poly(xa, b, xa + c, b + s * dh, xb - c, b + s * dh, xb, b);
            }
        }

        public override Model3 Model(P p)
        {
            var m = new Model3();
            m.Box(0, 0, 0, p["L"], p["H"], p["W"], top: false);
            m.Lid(p["L"], p["H"], p["W"], 35, p["TK"]);
            return m;
        }
    }

    public sealed class TuckReverse : TuckBase
    {
        public override string Id => "tuck-reverse";
        public override string Name => "Teskari qopqoq (RTE)";
        protected override int BottomLid => 0;
    }

    public sealed class TuckStraight : TuckBase
    {
        public override string Id => "tuck-straight";
        public override string Name => "To'g'ri qopqoq (STE)";
        protected override int BottomLid => 2;
    }

    public sealed class Gable : Template
    {
        public override string Id => "gable";
        public override string Name => "Ruchkali quti";
        public override string Group => "Tuck-end";
        public override Param[] Params => new[] { L(160), W(100), H(120), G(15),
            new Param("HH", "Ruchka", 60, "Ruchka qismining balandligi"),
            new Param("HW", "Teshik eni", 80, "Ruchka teshigi eni") };

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], hh = p["HH"], hw = Min(p["HW"], l - 20);
            var xs = Body(d, h, p["G"], l, w, l, w);
            SlotFlaps(d, xs, 0, -1, w / 2);
            double y1 = h + w / 2, y2 = y1 + hh, c = Min(hh / 3, 15);
            for (int i = 0; i < 4; i++)
            {
                double xa = xs[i], xb = xs[i + 1];
                d.Rect(xa, h, xb, y1);
                if (i % 2 == 0)
                {
                    d.Poly(xa, y1, xb, y1, xb, y2 - c, xb - c, y2, xa + c, y2, xa, y2 - c);
                    double cx = (xa + xb) / 2, cy = y1 + hh * 0.55, th = Min(18, hh * 0.35);
                    d.CutPoly(cx - hw / 2 + th / 2, cy - th / 2, cx + hw / 2 - th / 2, cy - th / 2, cx + hw / 2, cy,
                              cx + hw / 2 - th / 2, cy + th / 2, cx - hw / 2 + th / 2, cy + th / 2, cx - hw / 2, cy);
                }
                else
                {
                    double xm = (xa + xb) / 2;
                    d.Crease(xa, h, xm, y1);
                    d.Crease(xb, h, xm, y1);
                }
            }
        }

        public override Model3 Model(P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], hh = p["HH"];
            var m = new Model3();
            m.Box(0, 0, 0, l, h, w, top: false);
            V3 a = new(0, h, 0), b = new(l, h, 0), c = new(l, h, w), e = new(0, h, w);
            V3 ra = new(0, h + w / 2, w / 2), rb = new(l, h + w / 2, w / 2);
            m.Quad(a, ra, rb, b); m.Quad(e, c, rb, ra);
            m.Face(a, e, ra); m.Face(b, rb, c);
            m.Quad(ra, rb, new V3(l, h + w / 2 + hh, w / 2), new V3(0, h + w / 2 + hh, w / 2));
            return m;
        }
    }

    public sealed class Pillow : Template
    {
        public override string Id => "pillow";
        public override string Name => "Yostiqcha (pillow)";
        public override string Group => "Tuck-end";
        public override Param[] Params => new[] { L(160), W(90), G(15),
            new Param("D", "Qavariq", 25, "Uchlaridagi yoy balandligi") };

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], dd = Min(p["D"], w / 2);
            double g = p["G"], c = Min(g * 0.6, l / 4);
            d.Poly(0, 0, -g, c, -g, l - c, 0, l);
            for (int i = 0; i < 2; i++)
            {
                double xa = i * w, xb = xa + w;
                var lowTop = Arc(xa, xb, l, dd); var hiTop = Arc(xa, xb, l, 2 * dd);
                var lowBot = Arc(xa, xb, 0, -dd); var hiBot = Arc(xa, xb, 0, -2 * dd);
                // panel chegarasi yoy bo'ylab -> qopqoq bilan umumiy yoy bukish bo'ladi
                var panel = new List<(double, double)>(lowBot);
                for (int k = lowTop.Count - 1; k >= 0; k--) panel.Add(lowTop[k]);
                d.Poly(panel);
                d.Poly(Lens(lowTop, hiTop));
                d.Poly(Lens(lowBot, hiBot));
            }
        }

        static List<(double x, double y)> Lens(List<(double x, double y)> inner, List<(double x, double y)> outer)
        {
            var r = new List<(double, double)>(inner);
            for (int k = outer.Count - 2; k >= 1; k--) r.Add(outer[k]);
            return r;
        }

        public override Model3 Model(P p)
        {
            double l = p["L"], w = p["W"], th = p["D"] * 1.2;
            var m = new Model3();
            int n = 12;
            for (int i = 0; i < n; i++)
            {
                double t0 = (double)i / n, t1 = (double)(i + 1) / n;
                double z0 = th * System.Math.Sin(System.Math.PI * t0), z1 = th * System.Math.Sin(System.Math.PI * t1);
                double x0 = l * t0, x1 = l * t1;
                m.Quad(new V3(x0, 0, 0), new V3(x1, 0, 0), new V3(x1, 0, z1), new V3(x0, 0, z0));
                m.Quad(new V3(x0, w, 0), new V3(x0, w, z0), new V3(x1, w, z1), new V3(x1, w, 0));
                m.Quad(new V3(x0, 0, z0), new V3(x1, 0, z1), new V3(x1, w, z1), new V3(x0, w, z0));
                m.Quad(new V3(x0, 0, -z0), new V3(x0, w, -z0), new V3(x1, w, -z1), new V3(x1, 0, -z1));
                m.Quad(new V3(x0, 0, 0), new V3(x0, 0, -z0), new V3(x1, 0, -z1), new V3(x1, 0, 0));
                m.Quad(new V3(x0, w, 0), new V3(x1, w, 0), new V3(x1, w, -z1), new V3(x0, w, -z0));
            }
            return m;
        }
    }

    // ================================================================ Mailer

    public class Mailer : Template
    {
        public override string Id => "mailer";
        public override string Name => "Mailer (FEFCO 0427)";
        public override string Group => "Mailer";
        public override Param[] Params => Extra(new[] { L(250), W(180), H(60), T(2) });
        protected virtual Param[] Extra(Param[] ps) => ps;

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], t = p["T"];
            d.Rect(0, 0, l, w);
            d.Rect(0, -h, l, 0);
            d.Rect(0, -2 * h + t, l, -h);
            d.Rect(-h, 0, 0, w);
            d.Rect(-2 * h + t, 0, -h, w);
            d.Rect(l, 0, l + h, w);
            d.Rect(l + h, 0, l + 2 * h - t, w);
            d.Rect(0, w, l, w + h);
            d.Rect(0, w + h, l, 2 * w + h);
            double tc = Min(h * 0.4, l / 6);
            d.Poly(0, 2 * w + h, tc, 2 * w + 2 * h - t, l - tc, 2 * w + 2 * h - t, l, 2 * w + h);
            double fc = Min(h * 0.5, w / 4);
            foreach (var (x0, sx) in new[] { (0.0, -1.0), (l, 1.0) })
            {
                d.Poly(x0, w + h + t, x0 + sx * (h - t), w + h + t + fc, x0 + sx * (h - t), 2 * w + h - fc, x0, 2 * w + h - t);
                foreach (var (y0, sy) in new[] { (0.0, -1.0), (w, 1.0) })
                    d.Poly(x0, y0 + sy * 0.5, x0 + sx * (h - 2 * t), y0 + sy * (0.5 + fc * 0.6),
                           x0 + sx * (h - 2 * t), y0 + sy * (h - 1), x0, y0 + sy * h);
            }
        }

        public override Model3 Model(P p)
        {
            var m = new Model3();
            m.Box(0, 0, 0, p["L"], p["H"], p["W"], top: false);
            m.Lid(p["L"], p["H"], p["W"], 70, p["H"] - p["T"]);
            return m;
        }
    }

    public sealed class MailerWindow : Mailer
    {
        public override string Id => "mailer-window";
        public override string Name => "Mailer, oynali";
        protected override Param[] Extra(Param[] ps)
        {
            var r = new List<Param>(ps) { new Param("M", "Oyna chetidan", 35, "Qopqoq chetidan oynagacha masofa") };
            return r.ToArray();
        }

        public override void Build(Die d, P p)
        {
            base.Build(d, p);
            double l = p["L"], w = p["W"], h = p["H"], m = Min(p["M"], Min(l, w) / 2 - 5);
            d.CutRect(m, w + h + m, l - m, 2 * w + h - m);
        }
    }

    // ================================================================ Lotok

    public class Tray : Template
    {
        public override string Id => "tray";
        public override string Name => "Lotok";
        public override string Group => "Lotok";
        public override Param[] Params => new[] { L(200), W(150), H(50) };

        public override void Build(Die d, P p) => Build(d, 0, p["L"], p["W"], p["H"]);

        public static void Build(Die d, double ox, double l, double w, double h)
        {
            d.Rect(ox, 0, ox + l, w);
            d.Rect(ox, -h, ox + l, 0);
            d.Rect(ox, w, ox + l, w + h);
            d.Rect(ox - h, 0, ox, w);
            d.Rect(ox + l, 0, ox + l + h, w);
            double c = h * 0.25;
            foreach (var (x0, sx) in new[] { (ox, -1.0), (ox + l, 1.0) })
                foreach (var (y0, sy) in new[] { (0.0, -1.0), (w, 1.0) })
                {
                    d.Poly(x0, y0, x0, y0 + sy * h, x0 + sx * (h - 2), y0 + sy * (h - c), x0 + sx * (h - 2), y0 + sy * c);
                    d.Slit(x0, y0, x0 + sx * h, y0);
                }
        }

        public override Model3 Model(P p) { var m = new Model3(); m.Box(0, 0, 0, p["L"], p["H"], p["W"], top: false); return m; }
    }

    public sealed class TrayDouble : Template
    {
        public override string Id => "tray-double";
        public override string Name => "Lotok, ikki qavat devor";
        public override string Group => "Lotok";
        public override Param[] Params => new[] { L(250), W(180), H(50), T(2) };

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], t = p["T"], c = h * 0.25;
            d.Rect(0, 0, l, w);
            d.Rect(-h, 0, 0, w);
            d.Rect(l, 0, l + h, w);
            d.Rect(0, -h, l, 0); d.Rect(0, -2 * h + t, l, -h);
            d.Rect(0, w, l, w + h); d.Rect(0, w + h, l, w + 2 * h - t);
            // burchak qanotlari yon devorlarga ulangan, ikki qavat devor ichiga kiradi
            foreach (var (x0, sx) in new[] { (0.0, -1.0), (l, 1.0) })
                foreach (var (y0, sy) in new[] { (0.0, -1.0), (w, 1.0) })
                {
                    d.Poly(x0, y0, x0 + sx * h, y0, x0 + sx * (h - c), y0 + sy * (h - 2 * t), x0 + sx * c, y0 + sy * (h - 2 * t));
                    d.Slit(x0, y0, x0, y0 + sy * h);
                }
        }

        public override Model3 Model(P p) { var m = new Model3(); m.Box(0, 0, 0, p["L"], p["H"], p["W"], top: false); return m; }
    }

    public sealed class Telescope : Template
    {
        public override string Id => "telescope";
        public override string Name => "Taglik + qopqoq";
        public override string Group => "Lotok";
        public override Param[] Params => new[] { L(200), W(150), H(50),
            new Param("HL", "Qopqoq bal.", 30, "Qopqoq devori balandligi"),
            new Param("K", "Zazor", 2, "Qopqoq va taglik orasidagi bo'shliq", 0, 0.5) };

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], hl = p["HL"], k = p["K"];
            Tray.Build(d, 0, l, w, h);
            Tray.Build(d, l + 2 * h + 20 + hl, l + 2 * k, w + 2 * k, hl);
        }

        public override Model3 Model(P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], hl = p["HL"], k = p["K"];
            var m = new Model3();
            m.Box(0, 0, 0, l, h, w, top: false);
            m.Box(-k, h + h * 0.6, -k, l + 2 * k, hl, w + 2 * k, bottom: false);
            return m;
        }
    }

    public sealed class Sleeve : Template
    {
        public override string Id => "sleeve";
        public override string Name => "G'ilof (sleeve)";
        public override string Group => "Lotok";
        public override Param[] Params => new[] { L(200), W(150), H(50), G(15) };

        public override void Build(Die d, P p) => Body(d, p["L"], p["G"], p["W"], p["H"], p["W"], p["H"]);

        public override Model3 Model(P p)
        {
            var m = new Model3();
            m.Box(0, 0, 0, p["L"], p["H"], p["W"], ends: false);
            return m;
        }
    }

    public sealed class Drawer : Template
    {
        public override string Id => "drawer";
        public override string Name => "Tortma quti";
        public override string Group => "Lotok";
        public override Param[] Params => new[] { L(200), W(150), H(50), G(15),
            new Param("K", "Zazor", 2, "G'ilof va lotok orasidagi bo'shliq", 0, 0.5) };

        public override void Build(Die d, P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], k = p["K"];
            Tray.Build(d, 0, l, w, h);
            double ox = l + h + 30 + p["G"];
            double ww = w + 2 * k, hh = h + k;
            double c = Min(p["G"] * 0.6, l / 4);
            double[] xs = { 0, ww, ww + hh, 2 * ww + hh, 2 * ww + 2 * hh };
            d.Poly(ox, 0, ox - p["G"], c, ox - p["G"], l + k - c, ox, l + k);
            for (int i = 0; i < 4; i++) d.Rect(ox + xs[i], 0, ox + xs[i + 1], l + k);
        }

        public override Model3 Model(P p)
        {
            double l = p["L"], w = p["W"], h = p["H"], k = p["K"];
            var m = new Model3();
            m.Box(0, 0, 0, l, h, w, top: false);
            m.Box(l * 0.4, -k, -k, l + k, h + 2 * k, w + 2 * k, ends: false);
            return m;
        }
    }
}
