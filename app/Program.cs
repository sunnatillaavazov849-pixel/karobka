using System;
using System.Windows;
using System.Windows.Markup;

namespace Karobka
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            var app = new Application { ShutdownMode = ShutdownMode.OnMainWindowClose };
            using var s = typeof(Program).Assembly.GetManifestResourceStream("Main.xaml");
            var win = (Window)XamlReader.Load(s);
            new UI.MainController(win);
            app.Run(win);
        }
    }
}
