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
using System.Windows.Forms;
using System.IO;
using System.Configuration;

namespace FeeViewerPro
{
    /// <summary>
    /// Interaction logic for GeneralSettings.xaml
    /// </summary>
    public partial class GeneralSettings : Window
    {
        private string DefaultDbPath
        { get; set; }
        private string DefaultSrcPath
        { get; set; }

        public GeneralSettings()
        {
            InitializeComponent();
            DefaultDbPath = Properties.Settings.Default.DBDefaultPath;
            DefltDbPath.Text = DefaultDbPath;
            DefaultSrcPath = Properties.Settings.Default.ImportDefaultPath;
            DefltSrcPath.Text = DefaultSrcPath;
            DefaultZip.Text = Properties.Settings.Default.DefaultZipCode;
            this.FontFamily = new System.Windows.Media.FontFamily(Properties.Settings.Default.AppFontFamily);
            this.FontSize = Properties.Settings.Default.AppFontSize;
            label1.FontFamily = this.FontFamily;
            label1.FontSize = this.FontSize;
            label2.FontFamily = this.FontFamily;
            label2.FontSize = this.FontSize;
            label3.FontFamily = this.FontFamily;
            label3.FontSize = this.FontSize;
            label4.FontFamily = this.FontFamily;
            label4.FontSize = this.FontSize;
        }

        private void SrcChange_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog SrcFileDialog = new FolderBrowserDialog();
            if (SrcFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DefaultSrcPath = SrcFileDialog.SelectedPath;
                DefltSrcPath.Text = DefaultSrcPath;
            }
        }

        private void SettingsSave_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.DBDefaultPath = DefltDbPath.Text;
            Properties.Settings.Default.ImportDefaultPath = DefltSrcPath.Text;
            Properties.Settings.Default.DefaultZipCode = DefaultZip.Text;
            Properties.Settings.Default.UsePreviousDB = (bool)UsePreviousDB.IsChecked;
            Properties.Settings.Default.Save();
            App.myMainWin.ZipcodeIn.Text = DefaultZip.Text;
            App.myMainWin.AneStdZipcodeIn.Text = DefaultZip.Text;
            App.myMainWin.AneBaseZipcodeIn.Text = DefaultZip.Text;
            Close();
        }

        private void SettingsCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void DBChange_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog DbChgDialog = new FolderBrowserDialog();
            if (DbChgDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DefaultDbPath = DbChgDialog.SelectedPath;
                DefltDbPath.Text = DefaultDbPath;
            }

        }
    }
}
