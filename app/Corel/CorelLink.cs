using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Karobka.Geometry;

namespace Karobka.Corel
{
    /// <summary>Ishlab turgan CorelDRAW ga COM orqali ulanib chizadi</summary>
    public static class CorelLink
    {
        const int cdrMillimeter = 3;

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        static extern int CLSIDFromProgID(string progId, out Guid clsid);

        [DllImport("oleaut32.dll")]
        static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

        static dynamic App()
        {
            if (CLSIDFromProgID("CorelDRAW.Application", out var clsid) != 0)
                throw new InvalidOperationException("CorelDRAW topilmadi");
            if (GetActiveObject(ref clsid, IntPtr.Zero, out var obj) != 0 || obj == null)
                throw new InvalidOperationException("CorelDRAW ochilmagan");
            return obj;
        }

        public static void Draw(List<Seg> segs, string caption)
        {
            dynamic app = App();
            dynamic doc = app.ActiveDocument ?? app.CreateDocument();
            doc.Unit = cdrMillimeter;
            doc.BeginCommandGroup("Karobka");
            app.Optimization = true;
            try
            {
                dynamic layer = doc.ActiveLayer;
                dynamic sr = app.CreateShapeRange();

                dynamic cut = MakeCurve(app, doc, layer, segs, Kind.Cut);
                if (cut != null)
                {
                    cut.Name = "Kesish";
                    cut.Outline.Width = 0.25;
                    cut.Outline.Color.RGBAssign(0, 90, 255);
                    sr.Add(cut);
                }
                dynamic fold = MakeCurve(app, doc, layer, segs, Kind.Fold);
                if (fold != null)
                {
                    fold.Name = "Bukish";
                    fold.Outline.Width = 0.25;
                    fold.Outline.Color.RGBAssign(230, 0, 40);
                    try { fold.Outline.Style = app.OutlineStyles[3]; } catch { }
                    sr.Add(fold);
                }
                if (sr.Count == 0) return;

                dynamic txt = layer.CreateArtisticText(sr.LeftX, sr.BottomY - 8, caption);
                try { txt.Text.Story.Size = 10; } catch { }
                sr.Add(txt);

                dynamic g = sr.Group();
                g.Name = "Karobka";
                dynamic page = doc.ActivePage;
                g.CenterX = page.SizeWidth / 2;
                g.CenterY = page.SizeHeight / 2;
                try { g.CenterX = page.CenterX; g.CenterY = page.CenterY; } catch { }
                g.CreateSelection();
            }
            finally
            {
                app.Optimization = false;
                doc.EndCommandGroup();
                try { app.ActiveWindow.Refresh(); } catch { }
            }
        }

        static dynamic MakeCurve(dynamic app, dynamic doc, dynamic layer, List<Seg> segs, Kind kind)
        {
            var chains = Die.Chains(segs, kind);
            if (chains.Count == 0) return null;
            dynamic crv = app.CreateCurve(doc);
            foreach (var ch in chains)
            {
                dynamic sp = crv.CreateSubPath(ch[0].x, ch[0].y);
                bool closed = ch.Count > 2 && ch[0] == ch[ch.Count - 1];
                int last = closed ? ch.Count - 2 : ch.Count - 1;
                for (int i = 1; i <= last; i++) sp.AppendLineSegment(ch[i].x, ch[i].y);
                if (closed) sp.Closed = true;
            }
            return layer.CreateCurve(crv);
        }
    }
}
