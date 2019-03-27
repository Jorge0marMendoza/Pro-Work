using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Forms;

namespace FeeViewerPro
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Create the startup window
            if (App.myMainWin == null)
            {
                App.myMainWin = new MainWindow();
                // Do stuff here, e.g. to the window
                App.myMainWin.Title = "FeeViewer Pro 3.2.0.1";
                // Show the window

                System.Drawing.Rectangle resolution = Screen.PrimaryScreen.Bounds;
                App.myMainWin.MaxWidth = resolution.Width - 20;
                App.myMainWin.MaxHeight = resolution.Height - 50;
                App.myMainWin.DGridSclrUCR.MaxHeight = resolution.Height - 400;
                App.myMainWin.dGridScrlAne.MaxHeight = resolution.Height - 400;
            }

            App.myMainWin.Show();
            App.myMainWin.SetVisibilityUCR();
            App.myMainWin.SetVisibilityAneStd();
            App.myMainWin.InvalidateVisual();
        }

        public static MainWindow myMainWin
        {
            get;
            private set;
        }
    }
}
