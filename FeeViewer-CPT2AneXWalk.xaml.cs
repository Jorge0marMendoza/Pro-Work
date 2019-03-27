using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
//using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows;
using System;

namespace FeeViewerPro
{
    /// <summary>
    /// Interaction logic for CPT2ANEXWALK.xaml
    /// </summary>
    public partial class CPT2ANEXWALK : Window
    {
        private DataTable AneData = new DataTable();
        public string sAneOut { get; set; }

        public CPT2ANEXWALK(string CPTIn)
        {
            InitializeComponent();
            AddToListbox(CPTIn);
        }

 
        private void AneSelected_Changed(object sender, SelectionChangedEventArgs e)
        {
            int nSelected;
            nSelected = AneListBox.SelectedIndex;
            sAneOut = (string)AneData.Rows[nSelected]["ANE"];
        }

        public void AddToListbox(string sCPT)
        {
            string sAne, sDescription;
            OleDbConnection connANEBase = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;
            int NumRows;

            string sSQL = "SELECT ACRX.ANE, ACPT.Description FROM ACRX INNER JOIN ACPT ON ACRX.ANE = ACPT.CPT where ACRX.ANE = ACPT.CPT and ACRX.cpt = '" + sCPT + "';";

            sAne = "";
            sDescription = "";
            try
            {
                connANEBase = new OleDbConnection(connectionString);
                OleDbCommand myAccessCmdFCTX = new OleDbCommand(sSQL, connANEBase);
                connANEBase.Open();

                var da = new OleDbDataAdapter(myAccessCmdFCTX);
                NumRows = da.Fill(AneData);
                if (NumRows < 1)
                {
                    MessageBox.Show("Unable to get the list of Anesthesia codes for the specified procedure code");
                    return;
                }

                connANEBase.Close();

                for (int i = 0; i < NumRows; i++)
                {
                    sAne = (string)AneData.Rows[i]["ANE"];
                    sDescription = (string)AneData.Rows[i]["Description"];

                    AneListBox.Items.Add(sAne + "\t" + sDescription);
                }
                AneListBox.SelectedIndex = 0;
                AneListBox.SelectedItem = 0;
                AneListBox.Focus();
            }
            catch (Exception Ex)
            {
                connANEBase.Close();
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            AneStd.returnAne = sAneOut;
            AneBase.returnAne = sAneOut;
            Close();
        }

        private void DoubleClickSelection(object sender, MouseButtonEventArgs e)
        {
            AneStd.returnAne = sAneOut;
            AneBase.returnAne = sAneOut;
            Close();
        }

        private void AneListDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AneStd.returnAne = sAneOut;
                AneBase.returnAne = sAneOut;
                Close();
            }

        }
    }

    public class AneListItem
    {
        public string AnethesiaCode { get; set; }
        public string AneDescription { get; set; }
        public AneListItem(string code, string desc)
        {
            this.AnethesiaCode = code;
            this.AneDescription = desc;
        }
    }
}
