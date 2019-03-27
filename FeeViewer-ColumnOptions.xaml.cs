using System;
using System.Collections.Generic;
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
    /// Interaction logic for ColumnOptions.xaml
    /// </summary>
    public partial class ColumnOptions : Window
    {
        public ColumnOptions()
        {
            InitializeComponent();
            this.FontFamily = new System.Windows.Media.FontFamily(Properties.Settings.Default.AppFontFamily);
            this.FontSize = Properties.Settings.Default.AppFontSize;
            label1.FontFamily = this.FontFamily;
            label1.FontSize = this.FontSize;
            if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                checkBox25th.IsChecked = true;
            else
                checkBox25th.IsChecked = false;

            if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                checkBox30th.IsChecked = true;
            else
                checkBox30th.IsChecked = false;

            if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                checkBox35th.IsChecked = true;
            else
                checkBox35th.IsChecked = false;

            if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                checkBox40th.IsChecked = true;
            else
                checkBox40th.IsChecked = false;

            if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                checkBox45th.IsChecked = true;
            else
                checkBox45th.IsChecked = false;

            if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                checkBox50th.IsChecked = true;
            else
                checkBox50th.IsChecked = false;

            if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                checkBox55th.IsChecked = true;
            else
                checkBox55th.IsChecked = false;

            if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                checkBox60th.IsChecked = true;
            else
                checkBox60th.IsChecked = false;

            if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                checkBox65th.IsChecked = true;
            else
                checkBox65th.IsChecked = false;

            if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                checkBox70th.IsChecked = true;
            else
                checkBox70th.IsChecked = false;

            if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                checkBox75th.IsChecked = true;
            else
                checkBox75th.IsChecked = false;

            if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                checkBox80th.IsChecked = true;
            else
                checkBox80th.IsChecked = false;

            if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                checkBox85th.IsChecked = true;
            else
                checkBox85th.IsChecked = false;

            if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                checkBox90th.IsChecked = true;
            else
                checkBox90th.IsChecked = false;

            if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                checkBox95th.IsChecked = true;
            else
                checkBox95th.IsChecked = false;

        }

        private void OptionCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OptionSave_Click(object sender, RoutedEventArgs e)
        {
            if (checkBox25th.IsChecked == true)
                App.myMainWin.bV25th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV25th = System.Windows.Visibility.Hidden;

            if (checkBox30th.IsChecked == true)
                App.myMainWin.bV30th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV30th = System.Windows.Visibility.Hidden;

            if (checkBox35th.IsChecked == true)
                App.myMainWin.bV35th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV35th = System.Windows.Visibility.Hidden;

            if (checkBox40th.IsChecked == true)
                App.myMainWin.bV40th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV40th = System.Windows.Visibility.Hidden;

            if (checkBox45th.IsChecked == true)
                App.myMainWin.bV45th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV45th = System.Windows.Visibility.Hidden;

            if (checkBox50th.IsChecked == true)
                App.myMainWin.bV50th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV50th = System.Windows.Visibility.Hidden;

            if (checkBox55th.IsChecked == true)
                App.myMainWin.bV55th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV55th = System.Windows.Visibility.Hidden;

            if (checkBox60th.IsChecked == true)
                App.myMainWin.bV60th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV60th = System.Windows.Visibility.Hidden;

            if (checkBox65th.IsChecked == true)
                App.myMainWin.bV65th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV65th = System.Windows.Visibility.Hidden;

            if (checkBox70th.IsChecked == true)
                App.myMainWin.bV70th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV70th = System.Windows.Visibility.Hidden;

            if (checkBox75th.IsChecked == true)
                App.myMainWin.bV75th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV75th = System.Windows.Visibility.Hidden;

            if (checkBox80th.IsChecked == true)
                App.myMainWin.bV80th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV80th = System.Windows.Visibility.Hidden;

            if (checkBox85th.IsChecked == true)
                App.myMainWin.bV85th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV85th = System.Windows.Visibility.Hidden;

            if (checkBox90th.IsChecked == true)
                App.myMainWin.bV90th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV90th = System.Windows.Visibility.Hidden;

            if (checkBox95th.IsChecked == true)
                App.myMainWin.bV95th = System.Windows.Visibility.Visible;
            else
                App.myMainWin.bV95th = System.Windows.Visibility.Hidden;
            App.myMainWin.InvalidateVisual();
            Close();
        }

        private void OptionReset_Click(object sender, RoutedEventArgs e)
        {
            checkBox25th.IsChecked = false;
            checkBox30th.IsChecked = false;
            checkBox35th.IsChecked = false;
            checkBox40th.IsChecked = false;
            checkBox45th.IsChecked = false;
            checkBox50th.IsChecked = true;
            checkBox55th.IsChecked = false;
            checkBox60th.IsChecked = true;
            checkBox65th.IsChecked = false;
            checkBox70th.IsChecked = true;
            checkBox75th.IsChecked = true;
            checkBox80th.IsChecked = true;
            checkBox85th.IsChecked = true;
            checkBox90th.IsChecked = true;
            checkBox95th.IsChecked = true;
        }
    }
}
