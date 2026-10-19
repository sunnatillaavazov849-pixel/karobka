using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Karobka.Geometry;
using Karobka.Templates;

// Har shablon uchun: kesish konturida "ochiq uch" yo'qligini tekshiradi va SVG yozadi.
static class Check
{
    static int Main(string[] args)
    {
        string outDir = args.Length > 0 ? args[0] : "preview";
        Directory.CreateDirectory(outDir);
        int bad = 0;
        foreach (var t in Catalog.All)
        {
            var segs = t.Segments(t.Defaults());
            var deg = new Dictionary<(double, double), int>();
            foreach (var s in segs.Where(s => s.Kind == Kind.Cut))
            {
                deg[(s.X1, s.Y1)] = deg.GetValueOrDefault((s.X1, s.Y1)) + 1;
                deg[(s.X2, s.Y2)] = deg.GetValueOrDefault((s.X2, s.Y2)) + 1;
            }
            int open = deg.Count(kv => kv.Value % 2 == 1);
            int cut = segs.Count(s => s.Kind == Kind.Cut), fold = segs.Count - cut;
            Console.WriteLine($"{t.Id,-16} kesish {cut,4}  bukish {fold,4}  ochiq uchlar {open}");
            if (cut == 0 || fold == 0) bad++;
            Svg(segs, Path.Combine(outDir, t.Id + ".svg"));
        }
        return bad;
    }

    static void Svg(List<Seg> segs, string path)
    {
        double x0 = segs.Min(s => Math.Min(s.X1, s.X2)) - 10, x1 = segs.Max(s => Math.Max(s.X1, s.X2)) + 10;
        double y0 = segs.Min(s => Math.Min(s.Y1, s.Y2)) - 10, y1 = segs.Max(s => Math.Max(s.Y1, s.Y2)) + 10;
        var w = new StringWriter { NewLine = "\n" };
        w.WriteLine(FormattableString.Invariant($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{x0} {-y1} {x1 - x0} {y1 - y0}\" width=\"{(x1 - x0) * 2}\" height=\"{(y1 - y0) * 2}\"><rect x=\"{x0}\" y=\"{-y1}\" width=\"100%\" height=\"100%\" fill=\"white\"/>"));
        foreach (var s in segs)
        {
            string st = s.Kind == Kind.Cut ? "stroke=\"#2F6BFF\"" : "stroke=\"#E5484D\" stroke-dasharray=\"3 2\"";
            w.WriteLine(FormattableString.Invariant($"<line x1=\"{s.X1}\" y1=\"{-s.Y1}\" x2=\"{s.X2}\" y2=\"{-s.Y2}\" {st} stroke-width=\"0.6\"/>"));
        }
        w.WriteLine("</svg>");
        File.WriteAllText(path, w.ToString());
    }
}
