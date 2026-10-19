using System;
using System.Collections.Generic;

namespace Karobka.Geometry
{
    public enum Kind { Cut, Fold }

    public readonly struct Seg
    {
        public readonly Kind Kind;
        public readonly double X1, Y1, X2, Y2;
        public Seg(Kind k, double x1, double y1, double x2, double y2) { Kind = k; X1 = x1; Y1 = y1; X2 = x2; Y2 = y2; }
    }

    /// <summary>
    /// Razvyortka yadrosi. Shablon panellarni (ko'pburchak) qo'shadi:
    /// ikki panelga umumiy qirra -> bukish, bitta panelga tegishli qirra -> kesish,
    /// Slit bilan ustma-ust qirra -> kesish. Cut/Crease erkin chiziqlar (oyna, teshik, qo'shimcha bukish).
    /// </summary>
    public sealed class Die
    {
        const double Eps = 1e-6, Tol = 1e-3;

        readonly List<(double x1, double y1, double x2, double y2)> edges = new();
        readonly List<(double x1, double y1, double x2, double y2)> slits = new();
        readonly List<Seg> free = new();

        static double R(double v) => Math.Round(v, 3);

        public void Poly(params double[] v)
        {
            int n = v.Length / 2;
            for (int i = 0; i < n; i++)
            {
                int j = (i + 1) % n;
                edges.Add((R(v[2 * i]), R(v[2 * i + 1]), R(v[2 * j]), R(v[2 * j + 1])));
            }
        }

        public void Poly(IList<(double x, double y)> pts)
        {
            var v = new double[pts.Count * 2];
            for (int i = 0; i < pts.Count; i++) { v[2 * i] = pts[i].x; v[2 * i + 1] = pts[i].y; }
            Poly(v);
        }

        public void Rect(double x1, double y1, double x2, double y2) => Poly(x1, y1, x2, y1, x2, y2, x1, y2);

        public void Slit(double x1, double y1, double x2, double y2) => slits.Add((R(x1), R(y1), R(x2), R(y2)));

        /// <summary>Yopiq kesish konturi (oyna, teshik)</summary>
        public void CutPoly(params double[] v)
        {
            int n = v.Length / 2;
            for (int i = 0; i < n; i++)
            {
                int j = (i + 1) % n;
                free.Add(new Seg(Kind.Cut, R(v[2 * i]), R(v[2 * i + 1]), R(v[2 * j]), R(v[2 * j + 1])));
            }
        }

        public void CutRect(double x1, double y1, double x2, double y2) => CutPoly(x1, y1, x2, y1, x2, y2, x1, y2);

        public void Crease(double x1, double y1, double x2, double y2) => free.Add(new Seg(Kind.Fold, R(x1), R(y1), R(x2), R(y2)));

        static bool OnSeg(double px, double py, (double x1, double y1, double x2, double y2) s)
        {
            double dx = s.x2 - s.x1, dy = s.y2 - s.y1, ln2 = dx * dx + dy * dy;
            if (ln2 < Eps) return false;
            double t = ((px - s.x1) * dx + (py - s.y1) * dy) / ln2;
            if (t < -Eps || t > 1 + Eps) return false;
            return Math.Abs(s.x1 + t * dx - px) < Tol && Math.Abs(s.y1 + t * dy - py) < Tol;
        }

        public List<Seg> Solve()
        {
            var pts = new HashSet<(double, double)>();
            foreach (var e in edges) { pts.Add((e.x1, e.y1)); pts.Add((e.x2, e.y2)); }
            var count = new Dictionary<(double, double, double, double), int>();
            var order = new List<(double, double, double, double)>();
            var ts = new List<double>();
            foreach (var (x1, y1, x2, y2) in edges)
            {
                double dx = x2 - x1, dy = y2 - y1, ln2 = dx * dx + dy * dy;
                if (ln2 < Eps) continue;
                ts.Clear(); ts.Add(0); ts.Add(1);
                foreach (var (px, py) in pts)
                {
                    double t = ((px - x1) * dx + (py - y1) * dy) / ln2;
                    if (t > Eps && t < 1 - Eps && Math.Abs(x1 + t * dx - px) < Tol && Math.Abs(y1 + t * dy - py) < Tol)
                        ts.Add(t);
                }
                ts.Sort();
                for (int i = 0; i + 1 < ts.Count; i++)
                {
                    if (ts[i + 1] - ts[i] < Eps) continue;
                    var p = (R(x1 + ts[i] * dx), R(y1 + ts[i] * dy));
                    var q = (R(x1 + ts[i + 1] * dx), R(y1 + ts[i + 1] * dy));
                    var k = p.CompareTo(q) <= 0 ? (p.Item1, p.Item2, q.Item1, q.Item2) : (q.Item1, q.Item2, p.Item1, p.Item2);
                    if (count.TryGetValue(k, out int c)) count[k] = c + 1;
                    else { count[k] = 1; order.Add(k); }
                }
            }
            var res = new List<Seg>(order.Count + free.Count);
            foreach (var k in order)
            {
                var kind = count[k] >= 2 ? Kind.Fold : Kind.Cut;
                if (kind == Kind.Fold)
                {
                    double mx = (k.Item1 + k.Item3) / 2, my = (k.Item2 + k.Item4) / 2;
                    foreach (var s in slits) if (OnSeg(mx, my, s)) { kind = Kind.Cut; break; }
                }
                res.Add(new Seg(kind, k.Item1, k.Item2, k.Item3, k.Item4));
            }
            res.AddRange(free);
            return res;
        }

        /// <summary>Bir turdagi ketma-ket segmentlarni uzluksiz chiziqlarga birlashtiradi (Corel va plotter uchun)</summary>
        public static List<List<(double x, double y)>> Chains(List<Seg> segs, Kind kind)
        {
            var left = new List<Seg>();
            foreach (var s in segs) if (s.Kind == kind) left.Add(s);
            var used = new bool[left.Count];
            var byPt = new Dictionary<(double, double), List<int>>();
            void Add((double, double) p, int i) { if (!byPt.TryGetValue(p, out var l)) byPt[p] = l = new List<int>(); l.Add(i); }
            for (int i = 0; i < left.Count; i++) { Add((left[i].X1, left[i].Y1), i); Add((left[i].X2, left[i].Y2), i); }
            var res = new List<List<(double, double)>>();
            for (int i = 0; i < left.Count; i++)
            {
                if (used[i]) continue;
                used[i] = true;
                var path = new LinkedList<(double, double)>();
                path.AddLast((left[i].X1, left[i].Y1)); path.AddLast((left[i].X2, left[i].Y2));
                for (int end = 0; end < 2; end++)
                {
                    while (true)
                    {
                        var tip = end == 0 ? path.Last.Value : path.First.Value;
                        int next = -1;
                        foreach (int j in byPt[tip]) if (!used[j]) { next = j; break; }
                        if (next < 0) break;
                        used[next] = true;
                        var s = left[next];
                        var other = (s.X1, s.Y1) == tip ? (s.X2, s.Y2) : (s.X1, s.Y1);
                        if (end == 0) path.AddLast(other); else path.AddFirst(other);
                    }
                }
                res.Add(new List<(double, double)>(path));
            }
            return res;
        }
    }
}
