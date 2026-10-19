using System;
using System.Collections.Generic;
using Karobka.Geometry;

namespace Karobka.Templates
{
    /// <summary>Bitta o'lcham maydoni</summary>
    public sealed class Param
    {
        public string Key, Label, Tip;
        public double Default, Min = 1, Step = 1;
        public Param(string key, string label, double def, string tip = null, double min = 1, double step = 1)
        { Key = key; Label = label; Default = def; Tip = tip; Min = min; Step = step; }
    }

    public sealed class P : Dictionary<string, double> { }

    public readonly struct V3
    {
        public readonly double X, Y, Z;
        public V3(double x, double y, double z) { X = x; Y = y; Z = z; }
    }

    /// <summary>Yig'ilgan qutining sodda 3D modeli: x = uzunlik, y = balandlik, z = eni</summary>
    public sealed class Model3
    {
        public readonly List<V3[]> Faces = new();

        public void Quad(V3 a, V3 b, V3 c, V3 d) => Faces.Add(new[] { a, b, c, d });
        public void Face(params V3[] pts) => Faces.Add(pts);

        /// <summary>Quti yon devorlari va tagi. top=true bo'lsa qopqog'i ham</summary>
        public void Box(double x, double y, double z, double L, double H, double W, bool top = true, bool bottom = true, bool ends = true)
        {
            V3 p(double a, double b, double c) => new(x + a, y + b, z + c);
            if (bottom) Quad(p(0, 0, 0), p(L, 0, 0), p(L, 0, W), p(0, 0, W));
            if (top) Quad(p(0, H, 0), p(0, H, W), p(L, H, W), p(L, H, 0));
            Quad(p(0, 0, W), p(L, 0, W), p(L, H, W), p(0, H, W));   // old
            Quad(p(0, 0, 0), p(0, H, 0), p(L, H, 0), p(L, 0, 0));   // orqa
            if (ends)
            {
                Quad(p(0, 0, 0), p(0, 0, W), p(0, H, W), p(0, H, 0));
                Quad(p(L, 0, 0), p(L, H, 0), p(L, H, W), p(L, 0, W));
            }
        }

        /// <summary>Orqa yuqori qirrada ochilgan qopqoq (angle gradus, 0 = yopiq)</summary>
        public void Lid(double L, double H, double W, double angle, double front = 0)
        {
            double a = angle * Math.PI / 180, c = Math.Cos(a), s = Math.Sin(a);
            V3 at(double x, double d, double up) => new(x, H + d * s + up * c, d * c - up * s);
            Quad(at(0, 0, 0), at(0, W, 0), at(L, W, 0), at(L, 0, 0));
            if (front > 0) Quad(at(0, W, 0), at(0, W, -front), at(L, W, -front), at(L, W, 0));
        }
    }

    public abstract class Template
    {
        public abstract string Id { get; }
        public abstract string Name { get; }
        public abstract string Group { get; }
        public abstract Param[] Params { get; }
        public abstract void Build(Die d, P p);
        public abstract Model3 Model(P p);

        public P Defaults()
        {
            var p = new P();
            foreach (var x in Params) p[x.Key] = x.Default;
            return p;
        }

        public List<Seg> Segments(P p)
        {
            var d = new Die();
            Build(d, p);
            return d.Solve();
        }

        // ---------- umumiy o'lcham maydonlari ----------
        protected static Param L(double v) => new("L", "Uzunlik", v, "Ichki uzunlik");
        protected static Param W(double v) => new("W", "Eni", v, "Ichki eni");
        protected static Param H(double v) => new("H", "Balandlik", v, "Ichki balandlik");
        protected static Param G(double v) => new("G", "Yelim", v, "Yelimlanadigan qanot eni");
        protected static Param T(double v) => new("T", "Qalinlik", v, "Karton qalinligi", 0.1, 0.5);
        protected static Param Tuck(double v) => new("TK", "Til", v, "Qopqoq tili (tuck) uzunligi");

        protected static double Min(double a, double b) => Math.Min(a, b);

        /// <summary>Yelim | w0 | w1 | ... yon devorlar. Panel chegaralari x larini qaytaradi</summary>
        protected static double[] Body(Die d, double H, double g, params double[] widths)
        {
            var xs = new double[widths.Length + 1];
            for (int i = 0; i < widths.Length; i++) xs[i + 1] = xs[i] + widths[i];
            double c = Min(g * 0.6, H / 4);
            d.Poly(0, 0, -g, c, -g, H - c, 0, H);
            for (int i = 0; i < widths.Length; i++) d.Rect(xs[i], 0, xs[i + 1], H);
            return xs;
        }

        /// <summary>FEFCO uslubidagi qanotlar (W/2), qo'shni qanotlar orasida tirqish</summary>
        protected static void SlotFlaps(Die d, double[] xs, double baseY, double s, double f)
        {
            for (int i = 0; i + 1 < xs.Length; i++)
                d.Rect(xs[i], baseY, xs[i + 1], baseY + s * f);
            for (int i = 1; i + 1 < xs.Length; i++)
                d.Slit(xs[i], baseY, xs[i], baseY + s * f);
        }

        /// <summary>a dan b gacha yoy (balandligi h), nuqtalar ro'yxati (uchlari bilan)</summary>
        protected static List<(double x, double y)> Arc(double xa, double xb, double y, double h, int n = 16)
        {
            var r = new List<(double, double)>();
            for (int i = 0; i <= n; i++)
            {
                double t = (double)i / n;
                r.Add((xa + (xb - xa) * t, y + h * Math.Sin(Math.PI * t)));
            }
            return r;
        }
    }
}
