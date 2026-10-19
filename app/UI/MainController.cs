using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Karobka.Corel;
using Karobka.Geometry;
using Karobka.Templates;

namespace Karobka.UI
{
    /// <summary>Main.xaml oynasini boshqaradi: shablon ro'yxati, o'lcham maydonlari, ko'rinish, chizish</summary>
    public sealed class MainController
    {
        readonly Window win;
        readonly StackPanel list, fields;
        readonly TextBox search;
        readonly TextBlock searchHint, title, status;
        readonly Canvas flat;
        readonly RadioButton modeFlat, mode3D;
        readonly FrameworkElement flatBox, view3DHost;
        readonly BoxView3D view3D;

        Template current;
        P values;
        readonly Dictionary<string, TextBox> boxes = new();
        readonly Dictionary<Template, Border> items = new();
        bool loading;

        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public MainController(Window w)
        {
            win = w;
            T F<T>(string n) where T : class => (T)w.FindName(n);
            list = F<StackPanel>("TemplateList");
            fields = F<StackPanel>("Fields");
            search = F<TextBox>("Search");
            searchHint = F<TextBlock>("SearchHint");
            title = F<TextBlock>("PreviewTitle");
            status = F<TextBlock>("Status");
            flat = F<Canvas>("Flat");
            modeFlat = F<RadioButton>("ModeFlat");
            mode3D = F<RadioButton>("Mode3D");
            flatBox = F<FrameworkElement>("FlatBox");
            view3DHost = F<FrameworkElement>("View3DHost");
            view3D = new BoxView3D(F<Viewport3D>("View3D"), (FrameworkElement)view3DHost);

            search.TextChanged += (_, _) => { searchHint.Visibility = search.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed; FillList(); };
            modeFlat.Checked += (_, _) => SetMode();
            mode3D.Checked += (_, _) => SetMode();
            F<Button>("DrawBtn").Click += (_, _) => DrawInCorel();

            FillList();
            Select(Catalog.All[0]);
        }

        Brush Res(string key) => (Brush)win.Resources[key];
        Style St(string key) => (Style)win.Resources[key];

        // ---------------------------------------------------------- shablon ro'yxati

        void FillList()
        {
            list.Children.Clear();
            items.Clear();
            string q = search.Text.Trim().ToLowerInvariant();
            string group = null;
            foreach (var t in Catalog.All)
            {
                if (q.Length > 0 && !(t.Name + " " + t.Group).ToLowerInvariant().Contains(q)) continue;
                if (t.Group != group)
                {
                    group = t.Group;
                    list.Children.Add(new TextBlock
                    {
                        Text = group, Style = St("Unit"),
                        Margin = new Thickness(10, list.Children.Count == 0 ? 0 : 12, 0, 4),
                    });
                }
                var b = new Border
                {
                    CornerRadius = new CornerRadius(10), Padding = new Thickness(10, 8, 10, 8), Cursor = Cursors.Hand,
                    Background = Brushes.Transparent,
                    Child = new TextBlock { Text = t.Name, Style = St("Label"), TextTrimming = TextTrimming.CharacterEllipsis },
                };
                var tt = t;
                b.MouseLeftButtonUp += (_, _) => Select(tt);
                b.MouseEnter += (_, _) => { if (tt != current) b.Background = Res("Field"); };
                b.MouseLeave += (_, _) => { if (tt != current) b.Background = Brushes.Transparent; };
                items[t] = b;
                list.Children.Add(b);
            }
            Highlight();
        }

        void Highlight()
        {
            foreach (var (t, b) in items)
            {
                bool on = t == current;
                b.Background = on ? Res("Field") : Brushes.Transparent;
                ((TextBlock)b.Child).FontWeight = on ? FontWeights.SemiBold : FontWeights.Normal;
            }
        }

        void Select(Template t)
        {
            current = t;
            values = t.Defaults();
            title.Text = t.Group + " · " + t.Name;
            BuildFields();
            Highlight();
            Refresh();
        }

        // ---------------------------------------------------------- o'lcham maydonlari

        void BuildFields()
        {
            loading = true;
            fields.Children.Clear();
            boxes.Clear();
            foreach (var p in current.Params)
            {
                var head = new DockPanel { Margin = new Thickness(2, 0, 2, 6), ToolTip = p.Tip };
                var unit = new TextBlock { Text = "mm", Style = St("Unit") };
                DockPanel.SetDock(unit, Dock.Right);
                head.Children.Add(unit);
                head.Children.Add(new TextBlock { Text = p.Label, Style = St("Label") });

                var box = new TextBox { Style = St("FieldBox"), TextAlignment = TextAlignment.Center, Text = Fmt(values[p.Key]) };
                var minus = new RepeatButton { Content = "−", Style = St("Step") };
                var plus = new RepeatButton { Content = "+", Style = St("Step") };
                var row = new Grid();
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.ColumnDefinitions.Add(new ColumnDefinition());
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                Grid.SetColumn(box, 1); Grid.SetColumn(plus, 2);
                row.Children.Add(minus); row.Children.Add(box); row.Children.Add(plus);
                var field = new Border { Background = Res("Field"), CornerRadius = new CornerRadius(10), Height = 36, Child = row };

                var pp = p;
                minus.Click += (_, _) => Bump(pp, -pp.Step);
                plus.Click += (_, _) => Bump(pp, pp.Step);
                box.TextChanged += (_, _) => OnEdit(pp, box);
                box.PreviewMouseWheel += (_, e) => { Bump(pp, e.Delta > 0 ? pp.Step : -pp.Step); e.Handled = true; };
                box.GotKeyboardFocus += (_, _) => box.SelectAll();

                boxes[p.Key] = box;
                fields.Children.Add(new StackPanel { Margin = new Thickness(0, 0, 0, 14), Children = { head, field } });
            }
            loading = false;
        }

        static string Fmt(double v) => v.ToString("0.##", Inv);

        static bool TryParse(string s, out double v) =>
            double.TryParse(s.Replace(',', '.').Trim(), NumberStyles.Float, Inv, out v);

        void Bump(Param p, double d)
        {
            double v = Math.Max(p.Min, Math.Round(values[p.Key] + d, 2));
            boxes[p.Key].Text = Fmt(v);
        }

        void OnEdit(Param p, TextBox box)
        {
            if (loading) return;
            if (TryParse(box.Text, out double v) && v >= p.Min)
            {
                box.Foreground = Res("Text");
                values[p.Key] = v;
                Refresh();
            }
            else
            {
                box.Foreground = Brushes.IndianRed;
                Status("Qiymat noto'g'ri: " + p.Label, true);
            }
        }

        // ---------------------------------------------------------- ko'rinish

        List<Seg> segs;

        void Refresh()
        {
            try
            {
                segs = current.Segments(values);
                DrawFlat();
                view3D.Show(current.Model(values));
                var (w, h) = Size(segs);
                Status($"Yoyilma: {w:0} × {h:0} mm", false);
            }
            catch (Exception ex)
            {
                segs = null;
                Status("Xato: " + ex.Message, true);
            }
        }

        static (double w, double h) Size(List<Seg> s)
        {
            double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
            foreach (var g in s)
            {
                x0 = Math.Min(x0, Math.Min(g.X1, g.X2)); x1 = Math.Max(x1, Math.Max(g.X1, g.X2));
                y0 = Math.Min(y0, Math.Min(g.Y1, g.Y2)); y1 = Math.Max(y1, Math.Max(g.Y1, g.Y2));
            }
            return (x1 - x0, y1 - y0);
        }

        void DrawFlat()
        {
            flat.Children.Clear();
            if (segs.Count == 0) return;
            double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
            foreach (var g in segs)
            {
                x0 = Math.Min(x0, Math.Min(g.X1, g.X2)); x1 = Math.Max(x1, Math.Max(g.X1, g.X2));
                y0 = Math.Min(y0, Math.Min(g.Y1, g.Y2)); y1 = Math.Max(y1, Math.Max(g.Y1, g.Y2));
            }
            double pad = Math.Max(x1 - x0, y1 - y0) * 0.03, sw = Math.Max(x1 - x0, y1 - y0) / 420;
            flat.Width = x1 - x0 + 2 * pad;
            flat.Height = y1 - y0 + 2 * pad;
            foreach (var kind in new[] { Kind.Fold, Kind.Cut })
            {
                var geo = new StreamGeometry();
                using (var c = geo.Open())
                    foreach (var ch in Die.Chains(segs, kind))
                    {
                        Point P((double x, double y) q) => new(q.x - x0 + pad, y1 - q.y + pad);
                        c.BeginFigure(P(ch[0]), false, false);
                        for (int i = 1; i < ch.Count; i++) c.LineTo(P(ch[i]), true, false);
                    }
                geo.Freeze();
                var path = new System.Windows.Shapes.Path
                {
                    Data = geo, StrokeThickness = sw * (kind == Kind.Cut ? 1.2 : 1),
                    Stroke = new SolidColorBrush(kind == Kind.Cut ? Color.FromRgb(0x2F, 0x6B, 0xFF) : Color.FromRgb(0xE5, 0x48, 0x4D)),
                    StrokeLineJoin = PenLineJoin.Round,
                };
                if (kind == Kind.Fold) path.StrokeDashArray = new DoubleCollection { 4, 3 };
                flat.Children.Add(path);
            }
        }

        void SetMode()
        {
            bool is3d = mode3D.IsChecked == true;
            flatBox.Visibility = is3d ? Visibility.Collapsed : Visibility.Visible;
            view3DHost.Visibility = is3d ? Visibility.Visible : Visibility.Collapsed;
        }

        void Status(string text, bool error)
        {
            status.Text = text;
            status.Foreground = error ? Brushes.IndianRed : Res("Muted");
        }

        // ---------------------------------------------------------- CorelDRAW

        void DrawInCorel()
        {
            if (segs == null) { Status("Qiymat noto'g'ri", true); return; }
            try
            {
                win.Cursor = Cursors.Wait;
                string cap = $"{current.Name}  {Fmt(values.GetValueOrDefault("L"))} × {Fmt(values.GetValueOrDefault("W"))}" +
                             (values.ContainsKey("H") ? $" × {Fmt(values["H"])}" : "") + " mm";
                CorelLink.Draw(segs, cap);
                Status("Chizildi", false);
            }
            catch (Exception ex)
            {
                string msg = ex is InvalidOperationException ? ex.Message : "CorelDRAW bilan bog'lanib bo'lmadi";
                Status(msg, true);
                MessageBox.Show(win, msg + (ex is InvalidOperationException ? "" : "\n\n" + ex.Message), "Karobka",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally { win.Cursor = null; }
        }
    }
}
