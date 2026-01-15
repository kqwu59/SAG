using System;
using System.Windows;

namespace NettoieXLSX.V16
{
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            var app = new Application();
            var window = new MainWindow();
            app.Run(window);
        }
    }
}
