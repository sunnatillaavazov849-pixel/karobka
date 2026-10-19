using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using Karobka.Templates;

namespace Karobka.UI
{
    /// <summary>Yig'ilgan qutining 3D ko'rinishi: sichqoncha bilan aylantirish, g'ildirak bilan yaqinlashtirish</summary>
    public sealed class BoxView3D
    {
        readonly Viewport3D vp;
        readonly ModelVisual3D visual = new();
        readonly AxisAngleRotation3D yaw = new(new Vector3D(0, 1, 0), -35);
        readonly AxisAngleRotation3D pitch = new(new Vector3D(1, 0, 0), 22);
        readonly PerspectiveCamera cam = new() { FieldOfView = 35 };
        double dist = 600, zoom = 1;
        Point? drag;

        static readonly Brush Kraft = new SolidColorBrush(Color.FromRgb(0xCF, 0xA9, 0x7E));
        static readonly Brush Inside = new SolidColorBrush(Color.FromRgb(0xB8, 0x92, 0x68));

        public BoxView3D(Viewport3D viewport, FrameworkElement host)
        {
            vp = viewport;
            vp.Camera = cam;
            vp.Children.Add(new ModelVisual3D
            {
                Content = new Model3DGroup
                {
                    Children =
                    {
                        new AmbientLight(Color.FromRgb(0x70, 0x70, 0x70)),
                        new DirectionalLight(Color.FromRgb(0xB0, 0xB0, 0xB0), new Vector3D(-0.4, -0.8, -0.6)),
                        new DirectionalLight(Color.FromRgb(0x40, 0x40, 0x40), new Vector3D(0.6, 0.3, 0.8)),
                    }
                }
            });
            var tg = new Transform3DGroup();
            tg.Children.Add(new RotateTransform3D(yaw));
            tg.Children.Add(new RotateTransform3D(pitch));
            visual.Transform = tg;
            vp.Children.Add(visual);

            host.MouseLeftButtonDown += (_, e) => { drag = e.GetPosition(host); host.CaptureMouse(); };
            host.MouseLeftButtonUp += (_, _) => { drag = null; host.ReleaseMouseCapture(); };
            host.MouseMove += (_, e) =>
            {
                if (drag == null) return;
                var p = e.GetPosition(host);
                yaw.Angle += (p.X - drag.Value.X) * 0.5;
                pitch.Angle = Math.Clamp(pitch.Angle + (p.Y - drag.Value.Y) * 0.5, -85, 85);
                drag = p;
            };
            host.MouseWheel += (_, e) => { zoom = Math.Clamp(zoom * (e.Delta > 0 ? 0.9 : 1.1), 0.4, 3); UpdateCamera(); };
        }

        public void Show(Model3 m)
        {
            double x0 = double.MaxValue, y0 = double.MaxValue, z0 = double.MaxValue;
            double x1 = double.MinValue, y1 = double.MinValue, z1 = double.MinValue;
            foreach (var f in m.Faces)
                foreach (var v in f)
                {
                    x0 = Math.Min(x0, v.X); x1 = Math.Max(x1, v.X);
                    y0 = Math.Min(y0, v.Y); y1 = Math.Max(y1, v.Y);
                    z0 = Math.Min(z0, v.Z); z1 = Math.Max(z1, v.Z);
                }
            double cx = (x0 + x1) / 2, cy = (y0 + y1) / 2, cz = (z0 + z1) / 2;
            double size = Math.Max(x1 - x0, Math.Max(y1 - y0, z1 - z0));

            var mesh = new MeshGeometry3D();
            var edges = new MeshGeometry3D();
            double ew = size / 400;
            foreach (var f in m.Faces)
            {
                int start = mesh.Positions.Count;
                foreach (var v in f) mesh.Positions.Add(new Point3D(v.X - cx, v.Y - cy, v.Z - cz));
                for (int i = 1; i + 1 < f.Length; i++)
                {
                    mesh.TriangleIndices.Add(start); mesh.TriangleIndices.Add(start + i); mesh.TriangleIndices.Add(start + i + 1);
                }
                for (int i = 0; i < f.Length; i++)
                {
                    var a = f[i]; var b = f[(i + 1) % f.Length];
                    AddEdge(edges, new Point3D(a.X - cx, a.Y - cy, a.Z - cz), new Point3D(b.X - cx, b.Y - cy, b.Z - cz), ew);
                }
            }
            var g = new Model3DGroup();
            g.Children.Add(new GeometryModel3D(mesh, new DiffuseMaterial(Kraft)) { BackMaterial = new DiffuseMaterial(Inside) });
            var edgeMat = new DiffuseMaterial(new SolidColorBrush(Color.FromRgb(0x6B, 0x4E, 0x2E)));
            g.Children.Add(new GeometryModel3D(edges, edgeMat) { BackMaterial = edgeMat });
            visual.Content = g;

            dist = size * 2.6;
            UpdateCamera();
        }

        /// <summary>Qirra chizig'i: ingichka to'rtburchak prizma</summary>
        static void AddEdge(MeshGeometry3D m, Point3D a, Point3D b, double w)
        {
            var d = b - a;
            if (d.Length < 1e-6) return;
            d.Normalize();
            var u = Vector3D.CrossProduct(d, Math.Abs(d.Y) < 0.9 ? new Vector3D(0, 1, 0) : new Vector3D(1, 0, 0));
            u.Normalize();
            var v = Vector3D.CrossProduct(d, u);
            u *= w; v *= w;
            int s = m.Positions.Count;
            foreach (var p in new[] { a, b })
            {
                m.Positions.Add(p + u + v); m.Positions.Add(p - u + v); m.Positions.Add(p - u - v); m.Positions.Add(p + u - v);
            }
            for (int i = 0; i < 4; i++)
            {
                int j = (i + 1) % 4;
                m.TriangleIndices.Add(s + i); m.TriangleIndices.Add(s + 4 + i); m.TriangleIndices.Add(s + 4 + j);
                m.TriangleIndices.Add(s + i); m.TriangleIndices.Add(s + 4 + j); m.TriangleIndices.Add(s + j);
            }
        }

        void UpdateCamera()
        {
            double d = dist * zoom;
            cam.Position = new Point3D(0, 0, d);
            cam.LookDirection = new Vector3D(0, 0, -d);
            cam.UpDirection = new Vector3D(0, 1, 0);
            cam.NearPlaneDistance = d / 100;
            cam.FarPlaneDistance = d * 10;
        }
    }
}
