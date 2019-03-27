using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FeeViewerPro
{
    /// <summary>
    /// Interaction logic for Splash.xaml
    /// </summary>
    public partial class Splash : Window
    {
        string DocFile;

        public Splash(string sDocfile)
        {
            DocFile = sDocfile;
            InitializeComponent();
        }

        private void Splash_Continue_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Doc_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(DocFile) == false)
            {
                System.Windows.MessageBox.Show("Can't find the user guide");
            }
            else
            {
                Process.Start(DocFile);
            }
        }
    }

}
