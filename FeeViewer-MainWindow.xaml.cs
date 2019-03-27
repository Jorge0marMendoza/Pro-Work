using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Windows.Documents;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;

namespace FeeViewerPro
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        // UCR tables
        public System.Data.DataTable UcrTable = new System.Data.DataTable();        // static table with column defines
        public System.Data.DataTable UCRData;                           // final table will all data - instance of UcrTable
        public System.Data.DataTable CodeTable;                           // final table will all data - instance of UcrTable
        public System.Data.DataTable CptFlagDescRvuModUcrTable;              // table with Code, Flag, Description, RVU and Modifier
        public System.Data.DataTable FctUcrTable;                            // Table containing FCT table query results
        // Ane Standard tables
        public System.Data.DataTable AneStdTable = new System.Data.DataTable();
        public System.Data.DataTable ANEStdData;
        public System.Data.DataTable CptFlagDescRvuModAneStd;
        //ANE Base tables
        public System.Data.DataTable AneBaseTable = new System.Data.DataTable();
        public System.Data.DataTable ANEBaseData;
        public System.Data.DataTable CptFlagDescModRvuAneBase;

        public event PropertyChangedEventHandler PropertyChanged;

        public System.Data.DataTable ThatProperty
        {
            get
            {
                return UCRData;
            }
            set
            {
                UCRData = value;
            }
        }

        public string SingleCpt
        {
            get
            {
                return CPTCode;
            }
        }

        public Visibility bV25th
        { get; set; }

        public Visibility bV30th
        { get; set; }

        public Visibility bV35th
        { get; set; }

        public Visibility bV40th
        { get; set; }

        public Visibility bV45th
        { get; set; }

        public Visibility bV50th
        { get; set; }

        public Visibility bV55th
        { get; set; }

        public Visibility bV60th
        { get; set; }

        public Visibility bV65th
        { get; set; }

        public Visibility bV70th
        { get; set; }

        public Visibility bV75th
        { get; set; }

        public Visibility bV80th
        { get; set; }

        public Visibility bV85th
        { get; set; }

        public Visibility bV90th
        { get; set; }

        public Visibility bV95th
        { get; set; }

        public bool bHasCPT
        { get; set; }

        public bool bHasAneCPT
        { get; set; }

        public bool bHasDRG
        { get; set; }

        public string CPTCode;
        public string Modifier;
        public string ZipCode;

        private string cptstart;
        private string cptend;

        private bool bIsRange;
        public int RVSID;
        public string GEOZIP;

        private string sMrufile1;
        private string sMrufile2;

        public bool bIsInitUCR = true;
        public bool bIsInitANEStd = true;
        public bool bIsInitANEBase = true;

        public string sImportWorkingPath = string.Empty;
        public string sDocFile;

        //public List<string> ModifierList = new List<string>();
        public List<string> ModifierList1 { get; set; }
        public List<string> ModifierList2 { get; set; }

        public string DocFile
        {
            get;
            set;
        }

        public MainWindow()
        {
            string sDocPath = System.IO.Path.GetDirectoryName( System.Reflection.Assembly.GetExecutingAssembly().Location);
            //sDocFile = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "FeeViewer 3.0 User Documentation.docx");
            sDocFile = System.IO.Path.Combine(sDocPath, "FeeViewer 3.2 Help Menu.chm");
            string WordMissing;
            WordMissing = "All documention is provided in Microsoft Word DOCX format." + Environment.NewLine;
            WordMissing = WordMissing + "You will need software to view the documents." + Environment.NewLine;
            WordMissing = WordMissing + "Software other then Word may not display properly" + Environment.NewLine + Environment.NewLine;
            WordMissing = WordMissing + "This will try WordPad or Word Viewer if available.";
            try
            {
                Type officeType = Type.GetTypeFromProgID("Word.Application");
                if (officeType == null)
                {
                   // System.Windows.MessageBox.Show(WordMissing);
                }
                // This should handle both CD and local folder installs
                if (File.Exists(sDocFile) == false)
                {
                    sDocFile = "F:\\cd_images\\FeeViewerPro\\FeeViewer 3.2 Help Menu.chm";
                    if (File.Exists(sDocFile) == false)
                    {
                        System.Windows.MessageBox.Show("Can't find the user guide: " + sDocFile);
                    }
                }
                DocFile = sDocFile;
            }
            catch (SystemException ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Windows.MessageBox.Show("User Guide Missing");
            }

            Splash CPTCopyright = new Splash(sDocFile);
            CPTCopyright.ShowDialog();
            InitializeComponent();
            this.DataContext = this;

            // Settings stuff
            if (Properties.Settings.Default.UpgradeSettings == true)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeSettings = false;
                Properties.Settings.Default.Save();
                Properties.Settings.Default.Reload();
            }
            GetColumnSettings();
            SetUCRTabFonts();
            System.Drawing.Rectangle resolution = Screen.PrimaryScreen.Bounds;
            UCRGrid.MaxWidth = resolution.Width - 80;
            ANEStdDataGrid.MaxWidth = resolution.Width - 80;
            ANEBaseGrid1.MaxWidth = resolution.Width - 80;
            UCRData = new System.Data.DataTable();
            UCRTableNewColumns(UcrTable);

            UCRData = UcrTable;
            CodeTable = CptFlagDescRvuModUcrTable;

            System.Data.DataTable ANEStdData = new System.Data.DataTable();
            ANESTDTableNewColumns(AneStdTable);
            ANEStdData = AneStdTable;

            ANEBaseData = new System.Data.DataTable();
            ANEBaseTableNewColumns(AneBaseTable);
            ANEBaseData = AneBaseTable;

            bHasCPT = false;
            bHasAneCPT = false;
            bHasDRG = false;

            UCRGrid.ItemsSource = UCRData.DefaultView;

            ANEStdDataGrid.ItemsSource = ANEStdData.DefaultView;

            ANEBaseGrid1.ItemsSource = ANEBaseData.DefaultView;

            if (string.IsNullOrEmpty(Properties.Settings.Default.MRUFile1) != true)
            {
                sMrufile1 = Properties.Settings.Default.MRUFile1;
                MRUFL1.Header = sMrufile1;
                MRUFL1.Visibility = System.Windows.Visibility.Visible;
                MRUFL1.IsEnabled = true;
                Seperator2.Visibility = System.Windows.Visibility.Visible;
            }
            if (string.IsNullOrEmpty(Properties.Settings.Default.MRUFile2) != true)
            {
                sMrufile2 = Properties.Settings.Default.MRUFile2;
                MRUFL2.Header = sMrufile2;
                MRUFL2.Visibility = System.Windows.Visibility.Visible;
                MRUFL2.IsEnabled = true;
            }
            SearchClearDisable();
            ZipcodeIn.Text = Properties.Settings.Default.DefaultZipCode;
            AneStdZipcodeIn.Text = Properties.Settings.Default.DefaultZipCode;
            AneBaseZipcodeIn.Text = Properties.Settings.Default.DefaultZipCode;

            AneStdPSM.SelectedIndex = 0;
            AneBasePSM.SelectedIndex = 0;
            int CurrentYear = DateTime.Now.Year;
            int CurrentMonth = DateTime.Now.Month;
            int CPTYear;
            //If The date is Last month of year (December) Increase copyright years by 1
            if (CurrentMonth == 12)
            {
                CPTYear = CurrentYear;
                CurrentYear = CurrentYear + 1;
            }
            else
                CPTYear = CurrentYear - 1;

            CopyrightNotice.Text = "Copyright® " + CurrentYear + " Context4 Healthcare, Inc." + Environment.NewLine + "CPT® copyright " + CPTYear + " American Medical Association";
            ModifierList1 = new List<string>();
            ModifierList1.Add("");
            ModifierList1.Add("26");
            ModifierList2 = new List<string>();
            ModifierList2.Add("");
            ModifierList2.Add("26");
            ModifierList2.Add("NU");
            ModifierList2.Add("RR");
            ModifierList2.Add("UE");
            ModifierIn.ItemsSource = ModifierList1;
            if (Properties.Settings.Default.UsePreviousDB == true)
            {
                MRU1_Open();
            }
            sImportWorkingPath = Properties.Settings.Default.ImportDefaultPath;
        }

        private void SearchClearDisable()
        {
            Search.IsEnabled = false;
            AneBaseSearch.IsEnabled = false;
            AneStdSearch.IsEnabled = false;
            Clear.IsEnabled = false;
            AneStdClear.IsEnabled = false;
            AneBaseClear.IsEnabled = false;
        }

        private void SearchClearEnable()
        {
            Search.IsEnabled = true;
            AneBaseSearch.IsEnabled = true;
            AneStdSearch.IsEnabled = true;
            Clear.IsEnabled = true;
            AneStdClear.IsEnabled = true;
            AneBaseClear.IsEnabled = true;
        }

        private void AppClose_Click(object sender, RoutedEventArgs e)
        {
            SaveColumnSettings();
            Close();
        }

        private void Search_Click(object sender, RoutedEventArgs e)
        {
            /* RVSID is determined by radio button selection
             * 777 - Medical
             * 776 - Dental
             * 775 - HCPCS
             * 774 - Out Patient
             * 773 - Inpatient by Patient
             * 772 - Inpatient by Day
            */
            bool bIsNum, bSuccess;
            int testzip, RVSID;
            bool bValidZip;
            string Modifier, GEOZIP, CPTCode, SelectType, Status;

            GEOZIP = "";
            Status = "";
            Modifier = ModifierIn.Text;
            if (((bIsNum = int.TryParse(ZipcodeIn.Text, out testzip)) != true) || (ZipcodeIn.Text.Length != 5))
            {
                System.Windows.MessageBox.Show("Please enter a valid Zipcode!");
                return;
            }
            else
            {
                //bValidZip = Common.ValidateZip(testzip, ref GEOZIP);
                bValidZip = Common.ValidateZip(ZipcodeIn.Text, ref GEOZIP);
                if (bValidZip == false)
                    return;
            }
            CPTCode = CptcodeIn.Text;
            if (CPTCode.Length == 11)
            {
                cptstart = CPTCode.Substring(0, 5);
                cptend = CPTCode.Substring(CPTCode.Length - 5, 5);
                bIsRange = true;
            }
            else if (CPTCode.Length == 5)
            {
                cptstart = CPTCode;
                cptend = CPTCode;
                bIsRange = false;
            }
            else
            {
                System.Windows.MessageBox.Show("Enter a valid code (length=5)");
                return;
            }

            if (ButtonMedical.IsChecked == true)        //Medical
            {
                RVSID = 777;
                SelectType = "MED";
            }
            else if (ButtonDental.IsChecked == true)    //Dental
            {
                RVSID = 776;
                SelectType = "DEN";
            }
            else if (ButtonHCPCS.IsChecked == true)     //HCPCS
            {
                RVSID = 775;
                SelectType = "HCP";
            }
            else if (ButtonOP.IsChecked == true)        //OUTPATIENT
            {
                RVSID = 774;
                SelectType = "OUT";
            }
            else if (ButtonIPD.IsChecked == true)       //IP Day
            {
                RVSID = 772;
                SelectType = "IPD";                     //In Patient by day
            }
            else if (ButtonIPP.IsChecked == true)       //IP Patient
            {
                RVSID = 773;
                SelectType = "IPP";                     //In Patient by Patient
            }
            else
            {
                System.Windows.MessageBox.Show("No code category type has been selected");
                RVSID = 000;
                SelectType = "XXX";                     //In Patient by day
            }


            if ((bIsRange == false) && (Common.ValidateCPT(cptstart, Status, true) == false))
            {
                return;
            }

            if (bIsInitUCR)
            {
                bIsInitUCR = false;
                UcrTable.Clear();
            }

            using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
            {
                //bSuccess = DBAccess.getUCRdata(UcrTable, testzip, cptstart, RVSID, Modifier, GEOZIP, SelectType, bIsRange, cptend);
                bSuccess = DBAccess.getUCRdata(UcrTable, ZipcodeIn.Text, cptstart, RVSID, Modifier, GEOZIP, SelectType, bIsRange, cptend);
                UcrTotalCount.Text = this.UCRGrid.Items.Count.ToString();
                UCRGrid.ItemsSource = UCRData.DefaultView;
                if (this.UCRGrid.Items.Count > 0)
                {
                    EditClearRows.IsEnabled = true;
                    EditSelectAll.IsEnabled = true;
                    EditCopy.IsEnabled = true;
                    EditExport.IsEnabled = true;
                    UCRGrid.ScrollIntoView(UCRGrid.Items[UCRGrid.Items.Count - 1]);
                    UCRGrid.UpdateLayout();
                }
                else
                {
                    EditClearRows.IsEnabled = false;
                    EditSelectAll.IsEnabled = false;
                    EditCopy.IsEnabled = false;
                    EditExport.IsEnabled = false;
                }
            }
        }

        private static System.Data.DataTable UCRTableNewColumns(DataTable UcrTable)
        {

            int i;
            try
            {
                UcrTable.Columns.Add("Zipcode", typeof(string));
                UcrTable.Columns.Add("Code", typeof(string));
                UcrTable.Columns.Add("Status", typeof(string));
                UcrTable.Columns.Add("Description", typeof(string));
                UcrTable.Columns.Add("Modifier", typeof(string));
                UcrTable.Columns.Add("Type", typeof(string));
                UcrTable.Columns.Add("25th", typeof(string));   //[5]
                UcrTable.Columns.Add("30th", typeof(string));   //[6]
                UcrTable.Columns.Add("35th", typeof(string));   //[7]
                UcrTable.Columns.Add("40th", typeof(string));   //[8]
                UcrTable.Columns.Add("45th", typeof(string));   //[9]
                UcrTable.Columns.Add("50th", typeof(string));   //[10]
                UcrTable.Columns.Add("55th", typeof(string));   //[11]
                UcrTable.Columns.Add("60th", typeof(string));   //[12]
                UcrTable.Columns.Add("65th", typeof(string));   //[13]
                UcrTable.Columns.Add("70th", typeof(string));   //[14]
                UcrTable.Columns.Add("75th", typeof(string));   //[15]
                UcrTable.Columns.Add("80th", typeof(string));   //[16]
                UcrTable.Columns.Add("85th", typeof(string));   //[17]
                UcrTable.Columns.Add("90th", typeof(string));   //[18]
                UcrTable.Columns.Add("95th", typeof(string));   //[19]
                DataRow dr = UcrTable.NewRow();
                dr["Description"] = "                              ";
                dr["25th"] = "     ";
                dr["30th"] = "     ";
                dr["35th"] = "     ";
                dr["40th"] = "     ";
                dr["45th"] = "     ";
                dr["50th"] = "     ";
                dr["55th"] = "     ";
                dr["60th"] = "     ";
                dr["65th"] = "     ";
                dr["70th"] = "     ";
                dr["75th"] = "     ";
                dr["80th"] = "     ";
                dr["85th"] = "     ";
                dr["90th"] = "     ";
                dr["95th"] = "     ";
                UcrTable.Rows.Add(dr);
                for (i = 1; i < 10; i++)
                {
                    dr = UcrTable.NewRow();
                    dr["25th"] = "     ";
                    UcrTable.Rows.Add(dr);
                }

                return UcrTable;
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return null;
            }

        }

        private static System.Data.DataTable CptFlagDescRvuModUcrNewCol(DataTable CptFlagDescRvuMod)
        {
            CptFlagDescRvuMod.Columns.Add("Code", typeof(string));
            CptFlagDescRvuMod.Columns.Add("Flag", typeof(string));
            CptFlagDescRvuMod.Columns.Add("Description", typeof(string));
            CptFlagDescRvuMod.Columns.Add("RVU", typeof(string));
            return CptFlagDescRvuMod;
        }

        private static System.Data.DataTable ANESTDTableNewColumns(DataTable AneStdTable)
        {
            int i;
            try
            {
                AneStdTable.Columns.Add("Zipcode", typeof(string));
                AneStdTable.Columns.Add("Code", typeof(string));
                AneStdTable.Columns.Add("Type", typeof(string));
                AneStdTable.Columns.Add("ANE", typeof(string));
                AneStdTable.Columns.Add("Status", typeof(string));
                AneStdTable.Columns.Add("Description", typeof(string));
                AneStdTable.Columns.Add("Minutes", typeof(string));
                AneStdTable.Columns.Add("PSM", typeof(string));
                AneStdTable.Columns.Add("25th", typeof(string));   //[5]
                AneStdTable.Columns.Add("30th", typeof(string));   //[6]
                AneStdTable.Columns.Add("35th", typeof(string));   //[7]
                AneStdTable.Columns.Add("40th", typeof(string));   //[8]
                AneStdTable.Columns.Add("45th", typeof(string));   //[9]
                AneStdTable.Columns.Add("50th", typeof(string));   //[10]
                AneStdTable.Columns.Add("55th", typeof(string));   //[11]
                AneStdTable.Columns.Add("60th", typeof(string));   //[12]
                AneStdTable.Columns.Add("65th", typeof(string));   //[13]
                AneStdTable.Columns.Add("70th", typeof(string));   //[14]
                AneStdTable.Columns.Add("75th", typeof(string));   //[15]
                AneStdTable.Columns.Add("80th", typeof(string));   //[16]
                AneStdTable.Columns.Add("85th", typeof(string));   //[17]
                AneStdTable.Columns.Add("90th", typeof(string));   //[18]
                AneStdTable.Columns.Add("95th", typeof(string));   //[19]
                DataRow dr = AneStdTable.NewRow();
                dr["Description"] = "                              ";
                dr["25th"] = "     ";
                dr["30th"] = "     ";
                dr["35th"] = "     ";
                dr["40th"] = "     ";
                dr["45th"] = "     ";
                dr["50th"] = "     ";
                dr["55th"] = "     ";
                dr["60th"] = "     ";
                dr["65th"] = "     ";
                dr["70th"] = "     ";
                dr["75th"] = "     ";
                dr["80th"] = "     ";
                dr["85th"] = "     ";
                dr["90th"] = "     ";
                dr["95th"] = "     ";
                AneStdTable.Rows.Add(dr);
                for (i = 1; i < 10; i++)
                {
                    dr = AneStdTable.NewRow();
                    dr["25th"] = "     ";
                    AneStdTable.Rows.Add(dr);
                }
                return AneStdTable;
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return null;
            }

        }

        private static System.Data.DataTable ANEBaseTableNewColumns(DataTable AneBaseTable)
        {
            int i;
            try
            {
                AneBaseTable.Columns.Add("Zipcode", typeof(string));
                AneBaseTable.Columns.Add("Code", typeof(string));
                AneBaseTable.Columns.Add("Type", typeof(string));
                AneBaseTable.Columns.Add("ANE", typeof(string));
                AneBaseTable.Columns.Add("Status", typeof(string));
                AneBaseTable.Columns.Add("Description", typeof(string));
                AneBaseTable.Columns.Add("Minutes", typeof(string));
                AneBaseTable.Columns.Add("PSM", typeof(string));
                AneBaseTable.Columns.Add("High Base High Minute", typeof(string));   //High Base / High Minute
                AneBaseTable.Columns.Add("High Base Low Minute", typeof(string));   //High Base / Low Minute
                AneBaseTable.Columns.Add("Low Base High Minute", typeof(string));   //Low Base / High Minute
                AneBaseTable.Columns.Add("Low Base Low Minute", typeof(string));   //Low Base / Low Minute
                DataRow dr = AneBaseTable.NewRow();
                //dr["High Base High Minute"] = "0.00";
                //dr["High Base Low Minute"] = "0.00";
                //dr["Low Base High Minute"] = "0.00";
                //dr["Low Base Low Minute"] = "0.00";
                dr["Description"] = "                              ";
                dr[8] = "     ";
                dr[9] = "     ";
                dr[10] = "     ";
                dr[11] = "     ";
                AneBaseTable.Rows.Add(dr);
                for (i = 1; i < 10; i++)
                {
                    dr = AneBaseTable.NewRow();
                    dr["Zipcode"] = "     ";
                    AneBaseTable.Rows.Add(dr);
                }
                return AneBaseTable;
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return null;
            }

        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            string message = "This operation can NOT be undone, are you sure you want to clear all?";
            string caption = "Erase All Work";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            // Displays the MessageBox.
            result = System.Windows.Forms.MessageBox.Show(message, caption, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                int i;
                DataRow dr = UcrTable.NewRow();
                UcrTable.Clear();
                for (i = 0; i < 10; i++)
                {
                    dr = UcrTable.NewRow();
                    dr["25th"] = "     ";
                    UcrTable.Rows.Add(dr);
                }
                UcrSearchCount.Text = "0";
                UcrTotalCount.Text = "0";
                App.myMainWin.CptcodeIn.Text = "";
                App.myMainWin.ModifierIn.Text = "";
                App.myMainWin.CheckAllMods.IsChecked = false;
                App.myMainWin.InvalidateVisual();
            }
        }

        private void DoOptions_Click(object sender, RoutedEventArgs e)
        {
            ColumnOptions _getOptions = new ColumnOptions();
            _getOptions.ShowDialog();

            if (UCRGrid != null)
                UCRGrid.ItemsSource = UCRData.DefaultView;
            SetVisibilityUCR();
            if (ANEStdData != null)
                ANEStdDataGrid.ItemsSource = ANEStdData.DefaultView;
            SetVisibilityAneStd();
            App.myMainWin.InvalidateVisual();
        }

        public void SetVisibilityUCR()
        {
            if (this.UCRGrid.Columns.Count == 0)
                return;
            this.UCRGrid.Columns[0].Visibility = System.Windows.Visibility.Visible; //Zipcode
            this.UCRGrid.Columns[1].Visibility = System.Windows.Visibility.Visible; //Code
            this.UCRGrid.Columns[2].Visibility = System.Windows.Visibility.Visible; //Status
            this.UCRGrid.Columns[3].Visibility = System.Windows.Visibility.Visible; //Description
            this.UCRGrid.Columns[4].Visibility = System.Windows.Visibility.Visible; //Modifier
            this.UCRGrid.Columns[5].Visibility = System.Windows.Visibility.Visible; //Type
            this.UCRGrid.Columns[6].Visibility = bV25th;
            this.UCRGrid.Columns[7].Visibility = bV30th;
            this.UCRGrid.Columns[8].Visibility = bV35th;
            this.UCRGrid.Columns[9].Visibility = bV40th;
            this.UCRGrid.Columns[10].Visibility = bV45th;
            this.UCRGrid.Columns[11].Visibility = bV50th;
            this.UCRGrid.Columns[12].Visibility = bV55th;
            this.UCRGrid.Columns[13].Visibility = bV60th;
            this.UCRGrid.Columns[14].Visibility = bV65th;
            this.UCRGrid.Columns[15].Visibility = bV70th;
            this.UCRGrid.Columns[16].Visibility = bV75th;
            this.UCRGrid.Columns[17].Visibility = bV80th;
            this.UCRGrid.Columns[18].Visibility = bV85th;
            this.UCRGrid.Columns[19].Visibility = bV90th;
            this.UCRGrid.Columns[20].Visibility = bV95th;
        }

        public void SetVisibilityAneStd()
        {
            try
            {
                //System.Diagnostics.Debug.Print(this.ANEStdDataGrid.Columns.Count.ToString());
                if (this.ANEStdDataGrid.Columns.Count == 0)
                    return;

                this.ANEStdDataGrid.Columns[0].Visibility = System.Windows.Visibility.Visible; //Zipcode
                this.ANEStdDataGrid.Columns[1].Visibility = System.Windows.Visibility.Visible; //CPT
                this.ANEStdDataGrid.Columns[2].Visibility = System.Windows.Visibility.Visible; //Type
                this.ANEStdDataGrid.Columns[3].Visibility = System.Windows.Visibility.Visible; //ANE
                this.ANEStdDataGrid.Columns[4].Visibility = System.Windows.Visibility.Visible; //Status
                this.ANEStdDataGrid.Columns[5].Visibility = System.Windows.Visibility.Visible; //Description
                this.ANEStdDataGrid.Columns[6].Visibility = System.Windows.Visibility.Visible; //Minutes
                this.ANEStdDataGrid.Columns[7].Visibility = System.Windows.Visibility.Visible; //PSM
                this.ANEStdDataGrid.Columns[8].Visibility = bV25th;
                this.ANEStdDataGrid.Columns[9].Visibility = bV30th;
                this.ANEStdDataGrid.Columns[10].Visibility = bV35th;
                this.ANEStdDataGrid.Columns[11].Visibility = bV40th;
                this.ANEStdDataGrid.Columns[12].Visibility = bV45th;
                this.ANEStdDataGrid.Columns[13].Visibility = bV50th;
                this.ANEStdDataGrid.Columns[14].Visibility = bV55th;
                this.ANEStdDataGrid.Columns[15].Visibility = bV60th;
                this.ANEStdDataGrid.Columns[16].Visibility = bV65th;
                this.ANEStdDataGrid.Columns[17].Visibility = bV70th;
                this.ANEStdDataGrid.Columns[18].Visibility = bV75th;
                this.ANEStdDataGrid.Columns[19].Visibility = bV80th;
                this.ANEStdDataGrid.Columns[20].Visibility = bV85th;
                this.ANEStdDataGrid.Columns[21].Visibility = bV90th;
                this.ANEStdDataGrid.Columns[22].Visibility = bV95th;
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString());
                return;
            }
        }

        private void UCR_Select_Click(object sender, RoutedEventArgs e)
        {
            SetVisibilityUCR();
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            Stream myStream = null;
            string DBName, TempPath, InstallPath, newConnectString, mruName, sExt;
            string sStatus = string.Empty;

            OpenFileDialog DBDialog = new OpenFileDialog();
            newConnectString = string.Empty;
            DBName = string.Empty;

            // Open App.Config of executable
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            InstallPath = Properties.Settings.Default.DBDefaultPath;
            DBDialog.InitialDirectory = Properties.Settings.Default.DBDefaultPath;

            // Set filter options and filter index.
            DBDialog.Filter = "Access Files|*.mdb|All Files|*.*";
            DBDialog.FilterIndex = 1;
            DBDialog.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            if (DBDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    if ((myStream = DBDialog.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                        }
                        myStream.Close();
                        SearchClearEnable();
                        DBName = DBDialog.FileName;
                        TempPath = System.IO.Path.GetDirectoryName(DBName);
                        mruName = System.IO.Path.GetFileNameWithoutExtension(DBName);
                        sExt = System.IO.Path.GetExtension(DBName).ToUpper();
                        Properties.Settings.Default.DBPath = TempPath;
                        newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                        if (sExt != ".MDB")
                        {
                            System.Windows.MessageBox.Show("The Database must be an MDB file type");
                            return;
                        }

                        Properties.Settings.Default.DbConnectionString = newConnectString;
                        if (string.IsNullOrEmpty(mruName) == false)
                        {
                            MruListUpdate(DBName, "Open");
                        }
                        OpenDbName.Text = DBName;
                        OpenDbName.Foreground = System.Windows.Media.Brushes.Black;
                    }
                    if (Common.ValidateDB(newConnectString) == true)
                    {
                        SearchClearEnable();
                        newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                        Properties.Settings.Default.DbConnectionString = newConnectString;
                        OpenDbName.Text = DBName;
                        OpenDbName.Foreground = System.Windows.Media.Brushes.Black;
                        if (Common.ValidateCPT("12002", sStatus, false) == true)
                        {
                            bHasCPT = true;
                        }
                        else
                        {
                            bHasCPT = false;
                        }
                        if (Common.ValidateAneCPT("00600", sStatus, false) == true)
                        {
                            bHasAneCPT = true;
                        }
                        else
                        {
                            bHasAneCPT = false;
                        }
                        if (Common.ValidateDRG("00001", sStatus, false) == true)
                        {
                            bHasDRG = true;
                        }
                        else
                        {
                            bHasDRG = false;
                        }
                    }
                    DBAccess.CheckForInstalledModules(0);
                    sImportWorkingPath = Properties.Settings.Default.ImportDefaultPath;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void NewDB_Click(object sender, RoutedEventArgs e)
        {
            string sNewDbName, sExt, newConnectString;
            string TempPath;
            string mruName1;

            SaveFileDialog MyCreateNewDb = new SaveFileDialog();
            MyCreateNewDb.Filter = "Access Files|*.mdb|All Files|*.*";
            MyCreateNewDb.InitialDirectory = Properties.Settings.Default.DBDefaultPath;
            if (MyCreateNewDb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewDbName = MyCreateNewDb.FileName;
                sExt = System.IO.Path.GetExtension(sNewDbName).ToUpper();
                if (sExt != ".MDB")
                {
                    System.Windows.MessageBox.Show("The Database must be an MDB file type");
                    return;
                }

                if (File.Exists(sNewDbName) == true)
                {
                    System.Windows.MessageBox.Show("The database already exists! If you really want to use this name, delete the existing one first.");
                    return;
                }
                try
                {
                    string path = AppDomain.CurrentDomain.BaseDirectory;
                    string DBTempate = System.IO.Path.Combine(path, "FeeShell.mdb");

                    System.IO.File.Copy(DBTempate, sNewDbName);
                    if (File.Exists(sNewDbName) == false)
                        System.Windows.MessageBox.Show("Unable to verify new database was created");

                    newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sNewDbName;
                    Properties.Settings.Default.DbConnectionString = newConnectString;

                    DBAccess.CheckForInstalledModules(0);

                    TempPath = System.IO.Path.GetDirectoryName(sNewDbName);
                    Properties.Settings.Default.DBPath = TempPath;
                    mruName1 = System.IO.Path.GetFileNameWithoutExtension(sNewDbName);
                    if (string.IsNullOrEmpty(mruName1) == false)
                    {
                        MruListUpdate(sNewDbName, "New");
                    }
                    OpenDbName.Text = sNewDbName;
                    OpenDbName.Foreground = System.Windows.Media.Brushes.Black;

                    //just in case they already has a database "open"
                    bHasCPT = false;
                    bHasAneCPT = false;
                    bHasDRG = false;

                    SearchClearDisable();
                    sImportWorkingPath = Properties.Settings.Default.ImportDefaultPath;
                }
                catch (Exception Ex)
                {
                    System.Windows.MessageBox.Show(Ex.Message.ToString(), "Copy Template MDB Error");
                }
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            Version sVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string sDsiplayVersion = string.Format(@"{0}.{1}.{2}.{3}", sVersion.Major, sVersion.Minor, sVersion.Build, sVersion.Revision );
            System.Windows.MessageBox.Show("FeeViewer Pro Version: " + sDsiplayVersion);
        }

        private void MRU1_Click(object sender, RoutedEventArgs e)
        {
            string DBName, newConnectString, sStatus;
            sStatus = string.Empty;
            try
            {
                DBName = Properties.Settings.Default.mruFQName1;

                newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                if (Common.ValidateDB(newConnectString) == true)
                {
                    SearchClearEnable();
                    newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                    Properties.Settings.Default.DbConnectionString = newConnectString;
                    OpenDbName.Text = DBName;
                    OpenDbName.Foreground = System.Windows.Media.Brushes.Black;
                    if (Common.ValidateCPT("12002", sStatus, false) == true)
                    {
                        bHasCPT = true;
                    }
                    else
                    {
                        bHasCPT = false;
                    }
                    if (Common.ValidateAneCPT("00600", sStatus, false) == true)
                    {
                        bHasAneCPT = true;
                    }
                    else
                    {
                        bHasAneCPT = false;
                    }
                    if (Common.ValidateDRG("00001", sStatus, false) == true)
                    {
                        bHasDRG = true;
                    }
                    else
                    {
                        bHasDRG = false;
                    }
                    DBAccess.CheckForInstalledModules(0);
                }
                sImportWorkingPath = Properties.Settings.Default.ImportDefaultPath;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }
        }

        private void MRU2_Click(object sender, RoutedEventArgs e)
        {
            string DBName, newConnectString;
            string sStatus = string.Empty;

            try
            {
                DBName = Properties.Settings.Default.mruFQName2;
                newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                if (Common.ValidateDB(newConnectString) == true)
                {
                    SearchClearEnable();
                    newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                    Properties.Settings.Default.DbConnectionString = newConnectString;
                    OpenDbName.Text = DBName;
                    OpenDbName.Foreground = System.Windows.Media.Brushes.Black;
                    if (Common.ValidateCPT("12002", sStatus, false) == true)
                    {
                        bHasCPT = true;
                    }
                    else
                    {
                        bHasCPT = false;
                    }
                    if (Common.ValidateAneCPT("00600", sStatus, false) == true)
                    {
                        bHasAneCPT = true;
                    }
                    else
                    {
                        bHasAneCPT = false;
                    }
                    if (Common.ValidateDRG("00001", sStatus, false) == true)
                    {
                        bHasDRG = true;
                    }
                    else
                    {
                        bHasDRG = false;
                    }
                }
                DBAccess.CheckForInstalledModules(0);
                MruListUpdate(DBName, "MRU2");
                sImportWorkingPath = Properties.Settings.Default.ImportDefaultPath;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }
        }

        private void MRU1_Open()
        {
            string DBName, newConnectString, sStatus;
            sStatus = string.Empty;
            try
            {
                DBName = Properties.Settings.Default.mruFQName1;

                newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                if (Common.ValidateDB(newConnectString) == true)
                {
                    SearchClearEnable();
                    newConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName;
                    Properties.Settings.Default.DbConnectionString = newConnectString;
                    OpenDbName.Text = DBName;
                    OpenDbName.Foreground = System.Windows.Media.Brushes.Black;
                    if (Common.ValidateCPT("12002", sStatus, false) == true)
                    {
                        bHasCPT = true;
                    }
                    else
                    {
                        bHasCPT = false;
                    }
                    if (Common.ValidateAneCPT("00600", sStatus, false) == true)
                    {
                        bHasAneCPT = true;
                    }
                    else
                    {
                        bHasAneCPT = false;
                    }
                    if (Common.ValidateDRG("00001", sStatus, false) == true)
                    {
                        bHasDRG = true;
                    }
                    else
                    {
                        bHasDRG = false;
                    }
                }
                //DBAccess.CheckForInstalledModules(0);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }
        }

        private void Import_ANE(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;
            OpenFileDialog MyImportAne = new OpenFileDialog();

            sPath = sImportWorkingPath;

            MyImportAne.InitialDirectory = sPath;
            MyImportAne.FileName = "anecpt.txt";
            MyImportAne.Filter = "txt files (ANECPT.txt)|ANECPT.txt|All files (*.*)|*.*";
            if (MyImportAne.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportAne.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("anecpt.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the anecpt.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//anecrx.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//anezip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//anervs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//anexfcts.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//anexfcts.txt";   // Base unit FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;
                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertANE(sNewFileName);
                    DBAccess.CheckForInstalledModules(7);
                    System.Windows.MessageBox.Show("Anesthesia files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void Import_DNT(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;
            OpenFileDialog MyImportDnt = new OpenFileDialog();

            sPath = sImportWorkingPath;

            MyImportDnt.InitialDirectory = sPath;
            MyImportDnt.FileName = "dntcpt.txt";
            MyImportDnt.Filter = "txt files (dntcpt.txt)|dntcpt.txt|All files (*.*)|*.*";
            if (MyImportDnt.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportDnt.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("dntcpt.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the dntcpt.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//dntzip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//dntrvs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//dntxfct.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;

                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertDNT(sNewFileName);
                    DBAccess.CheckForInstalledModules(2);
                    System.Windows.MessageBox.Show("Dental files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void Import_HCP(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;

            OpenFileDialog MyImportHcp = new OpenFileDialog();
            sPath = sImportWorkingPath;

            MyImportHcp.InitialDirectory = sPath;
            MyImportHcp.FileName = "hpxcpt.txt";
            MyImportHcp.Filter = "txt files (hpxcpt.txt)|hpxcpt.txt|All files (*.*)|*.*";
            if (MyImportHcp.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportHcp.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("hpxcpt.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the hpxcpt.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//hpxzip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//hpxrvs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//hpxxfct.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;

                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertHCP(sNewFileName);
                    DBAccess.CheckForInstalledModules(3);
                    System.Windows.MessageBox.Show("HCPCS files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void Import_IPF_Day(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;

            OpenFileDialog MyImportIPD = new OpenFileDialog();
            sPath = sImportWorkingPath;

            MyImportIPD.InitialDirectory = sPath;
            MyImportIPD.FileName = "daydrg.txt";
            MyImportIPD.Filter = "txt files (daydrg.txt)|daydrg.txt|All files (*.*)|*.*";
            if (MyImportIPD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportIPD.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("daydrg.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the daydrg.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//dayzip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//dayrvs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//dayxfct.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;

                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertIPD(sNewFileName);
                    DBAccess.CheckForInstalledModules(5);
                    System.Windows.MessageBox.Show("Inpatient by Day files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void Import_IPF_Pat(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;

            OpenFileDialog MyImportIPP = new OpenFileDialog();
            sPath = sImportWorkingPath;

            MyImportIPP.InitialDirectory = sPath;
            MyImportIPP.FileName = "patdrg.txt";
            MyImportIPP.Filter = "txt files (patdrg.txt)|patdrg.txt|All files (*.*)|*.*";
            if (MyImportIPP.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportIPP.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("patdrg.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the patdrg.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//patzip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//patrvs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//patxfct.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;

                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertIPP(sNewFileName);
                    DBAccess.CheckForInstalledModules(6);
                    System.Windows.MessageBox.Show("Inpatient by Patient files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void Import_MED(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;

            OpenFileDialog MyImportCPT = new OpenFileDialog();
            sPath = sImportWorkingPath;

            MyImportCPT.InitialDirectory = sPath;
            MyImportCPT.FileName = "ucrcpt.txt";
            MyImportCPT.Filter = "txt files (ucrcpt.txt)|ucrcpt.txt|All files (*.*)|*.*";
            if (MyImportCPT.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportCPT.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("ucrcpt.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the ucrcpt.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//ucrzip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//ucrrvs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//ucrxfct.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;

                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertMED(sNewFileName);
                    DBAccess.CheckForInstalledModules(1);
                    System.Windows.MessageBox.Show("Medical files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void Import_OPF(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath, sIndvName;

            OpenFileDialog MyImportOPF = new OpenFileDialog();
            sPath = sImportWorkingPath;

            MyImportOPF.InitialDirectory = sPath;
            MyImportOPF.FileName = "opfcpt.txt";
            MyImportOPF.Filter = "txt files (opfcpt.txt)|opfcpt.txt|All files (*.*)|*.*";
            if (MyImportOPF.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sNewFileName = MyImportOPF.FileName.ToLower();
                sPath = System.IO.Path.GetDirectoryName(sNewFileName);
                if (sNewFileName.Contains("opfcpt.txt") == false)
                {
                    System.Windows.MessageBox.Show("You need to select the opfcpt.txt file");
                    return;
                }
                // Now test for the other files to be imported
                sIndvName = sPath + "//opfzip.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//opfrvs.txt";
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sIndvName = sPath + "//opfxfct.txt";   //Extended Standard FCT
                if (File.Exists(sIndvName) == false)
                {
                    System.Windows.MessageBox.Show("No Files Imported! Can't find file: " + sIndvName);
                    return;
                }
                sImportWorkingPath = sPath;

                using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
                {
                    M_Import.IsEnabled = false;
                    DBAccess.InsertOPF(sNewFileName);
                    DBAccess.CheckForInstalledModules(4);
                    System.Windows.MessageBox.Show("Outpatient files have been imported");
                    M_Import.IsEnabled = true;
                    SearchClearEnable();
                }
            }
        }

        private void MruListUpdate(string dbName, string whoCalled)
        {
            string mruName1, mruName2;  // MRU names - no extensions
            string tmpName;             // holds new MRU name

            tmpName = System.IO.Path.GetFileNameWithoutExtension(dbName);   //name no ext 
            mruName1 = Properties.Settings.Default.MRUFile1;                //get previous MRU1 name for compare
            mruName2 = Properties.Settings.Default.MRUFile2;                //get previous MRU2 name for compare

            if (whoCalled == "New")    //Move MRU1 to MRU2, then new MRU1
            {
                MRU1ToMRU2(tmpName, dbName);
            }
            else if (whoCalled == "Open")
            {

                if ((string.IsNullOrEmpty(mruName2) == false) && (tmpName == mruName2)) // Do first as it means there is an MRU1 defined
                {
                    SwitchMRUs();
                }
                else    // Shift MRUs and add new one
                {
                    MRU1ToMRU2(tmpName, dbName);
                }
            }
            else if (whoCalled == "MRU2")   // Switch MRU 1 and 2
            {
                SwitchMRUs();
            }
            Properties.Settings.Default.Save();
        }

        private void MRU1ToMRU2(string dbName, string fqName)
        {
            string tmpMRU;

            tmpMRU = Properties.Settings.Default.MRUFile1;
            if (string.IsNullOrEmpty(tmpMRU) == false)
            {
                // first set MRU2 to the old MRU1 if there was an MRU1
                Properties.Settings.Default.MRUFile2 = Properties.Settings.Default.MRUFile1;    // set new MRU2
                Properties.Settings.Default.mruFQName2 = Properties.Settings.Default.mruFQName1;   // set new FQMRU2
                MRUFL2.Header = Properties.Settings.Default.MRUFile1;
                MRUFL2.IsEnabled = true;
                MRUFL2.Visibility = System.Windows.Visibility.Visible;
                Seperator2.Visibility = System.Windows.Visibility.Visible;
                Seperator2.IsEnabled = true;
            }

            // now create the new MRU1
            Properties.Settings.Default.MRUFile1 = dbName;    // set new MRU2
            Properties.Settings.Default.mruFQName1 = fqName;   // set new FQMRU2
            MRUFL1.Header = dbName;
            MRUFL1.IsEnabled = true;
            MRUFL1.Visibility = System.Windows.Visibility.Visible;
            Seperator2.Visibility = System.Windows.Visibility.Visible;
            Seperator2.IsEnabled = true;
            Properties.Settings.Default.Save();
        }

        private void SwitchMRUs()
        {
            string mruName1, mruName2;  // MRU names - no extensions
            string fqName1, fqName2;             // MRU fully quallified name (with path) will get from MRU1 properties
            mruName1 = Properties.Settings.Default.MRUFile1;        //get previous MRU1 name to put in MRU2 (names - no ext)
            mruName2 = Properties.Settings.Default.MRUFile2;        //get previous MRU2 name to put in MRU1 (names - no ext)
            fqName1 = Properties.Settings.Default.mruFQName1;       //get previous FQMRU1 to put in FQMRU2 (fully quallified names)
            fqName2 = Properties.Settings.Default.mruFQName2;       //get previous FQMRU2 to put in FQMRU1 (fully quallified names)
            if (string.IsNullOrEmpty(mruName2) == false)            // there was an MRU2 to copy to MRU1
            {
                Properties.Settings.Default.MRUFile1 = mruName2;    // set new MRU1
                Properties.Settings.Default.mruFQName1 = fqName2;   // set new FQMRU1
                MRUFL1.Header = mruName2;
                MRUFL1.IsEnabled = true;
                MRUFL1.Visibility = System.Windows.Visibility.Visible;
            }
            Properties.Settings.Default.MRUFile2 = mruName1;        // set new MRU2
            Properties.Settings.Default.mruFQName2 = fqName1;    // set new FQMRU2
            MRUFL2.Header = mruName1;
            Properties.Settings.Default.Save();
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            GeneralSettings _getSettings = new GeneralSettings();
            _getSettings.ShowDialog();
        }

        private void AllMods_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAllMods.IsChecked == true)
            {
                ModifierIn.Text = "*";
                ModifierIn.IsEnabled = false;
            }
            else
            {
                if ((ButtonMedical.IsChecked == true) || (ButtonHCPCS.IsChecked == true))
                {
                    ModifierIn.IsEnabled = true;
                    ModifierIn.Text = "";
                }
                else
                {
                    ModifierIn.Text = "";
                    ModifierIn.IsEnabled = false;
                }
            }
        }

        private void SelectMed_Click(object sender, RoutedEventArgs e)
        {
            //ModifierList1.Clear();
            //ModifierList1.Add("");
            //ModifierList1.Add("26");
            App.myMainWin.ModifierIn.ItemsSource = ModifierList1;

            if (CheckAllMods.IsChecked == true)
            {
                ModifierIn.IsEnabled = false;
                ModifierIn.Text = "*";
            }
            else
            {
                ModifierIn.IsEnabled = true;
                ModifierIn.Text = "";
            }
            App.myMainWin.InvalidateVisual();
        }

        private void SelectDNT_Click(object sender, RoutedEventArgs e)
        {
            ModifierIn.IsEnabled = false;
            ModifierIn.Text = "";
        }

        private void SelectHCP_Click(object sender, RoutedEventArgs e)
        {
            App.myMainWin.ModifierIn.ItemsSource = ModifierList2;

            if (CheckAllMods.IsChecked == true)
            {
                ModifierIn.IsEnabled = false;
                ModifierIn.Text = "*";
            }
            else
            {
                ModifierIn.IsEnabled = true;
                ModifierIn.Text = "";
            }
            App.myMainWin.InvalidateVisual();
        }

        private void SelectIPD_Click(object sender, RoutedEventArgs e)
        {
            ModifierIn.IsEnabled = false;
            ModifierIn.Text = "";
        }

        private void SelectIPP_Click(object sender, RoutedEventArgs e)
        {
            ModifierIn.IsEnabled = false;
            ModifierIn.Text = "";
        }

        private void SelectOP_Click(object sender, RoutedEventArgs e)
        {
            ModifierIn.IsEnabled = false;
            ModifierIn.Text = "";
        }

        private void AneStdSearch_Click(object sender, RoutedEventArgs e)
        {
            string anecpt, aneminutes, anepsm, anezip;

            anecpt = AnestdCptcodeIn.Text;
            aneminutes = AneStdMinutes.Text;
            anepsm = AneStdPSM.Text;
            anezip = AneStdZipcodeIn.Text;
            if (bIsInitANEStd)
            {
                bIsInitANEStd = false;
                AneStdTable.Clear();
            }
            AneStd.AneStdSearch(anezip, anecpt, aneminutes, anepsm);
            if (this.ANEStdDataGrid.Items.Count > 0)
            {
                EditClearRows.IsEnabled = true;
                EditSelectAll.IsEnabled = true;
                EditCopy.IsEnabled = true;
                EditExport.IsEnabled = true;
            }
            else
            {
                EditClearRows.IsEnabled = false;
                EditSelectAll.IsEnabled = false;
                EditCopy.IsEnabled = false;
                EditExport.IsEnabled = false;
            }
        }

        private void AneStdClear_Click(object sender, RoutedEventArgs e)
        {
            string message = "This operation can NOT be undone, are you sure you want to clear all?";
            string caption = "Erase All Work";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            // Displays the MessageBox.
            result = System.Windows.Forms.MessageBox.Show(message, caption, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                int i;
                DataRow dr = AneStdTable.NewRow();
                AneStdTable.Clear();
                for (i = 0; i < 10; i++)
                {
                    dr = AneStdTable.NewRow();
                    dr["25th"] = "     ";
                    AneStdTable.Rows.Add(dr);
                }
                App.myMainWin.AnestdCptcodeIn.Text = "";
                App.myMainWin.AneStdMinutes.Text = "";
                App.myMainWin.InvalidateVisual();
            }
        }

        private void AneStd_Select_Click(object sender, MouseButtonEventArgs e)
        {
            SetVisibilityAneStd();
            App.myMainWin.InvalidateVisual();
        }

        private void AneBaseClear_Click(object sender, RoutedEventArgs e)
        {
             string message = "This operation can NOT be undone, are you sure you want to clear all?";
            string caption = "Erase All Work";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            // Displays the MessageBox.
            result = System.Windows.Forms.MessageBox.Show(message, caption, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {

                int i;
                DataRow dr = AneBaseTable.NewRow();
                AneBaseTable.Clear();
                for (i = 0; i < 10; i++)
                {
                    dr = AneBaseTable.NewRow();
                    dr["Zipcode"] = "     ";
                    AneBaseTable.Rows.Add(dr);
                }
                App.myMainWin.AneBaseCptcodeIn.Text = "";
                App.myMainWin.AneBaseMinutes.Text = "";
                bIsInitANEBase = true;
                App.myMainWin.InvalidateVisual();
            }

        }

        private void AneBaseSearch_Click(object sender, RoutedEventArgs e)
        {
            string anecpt, aneminutes, anepsm, anezip;

            anecpt = AneBaseCptcodeIn.Text;
            aneminutes = AneBaseMinutes.Text;
            anepsm = AneBasePSM.Text;
            anezip = AneBaseZipcodeIn.Text;
            if (bIsInitANEBase)
            {
                bIsInitANEBase = false;
                AneBaseTable.Clear();
            }
            AneBase.AneBaseSearch(anezip, anecpt, aneminutes, anepsm);
            if (this.ANEBaseGrid1.Items.Count > 0)
            {
                EditClearRows.IsEnabled = true;
                EditSelectAll.IsEnabled = true;
                EditCopy.IsEnabled = true;
                EditExport.IsEnabled = true;
            }
            else
            {
                EditClearRows.IsEnabled = false;
                EditSelectAll.IsEnabled = false;
                EditCopy.IsEnabled = false;
                EditExport.IsEnabled = false;
            }
        }

        private void UCRSelectAll_Click(object sender, RoutedEventArgs e)
        {
            UCRGrid.Focus();
            UCRGrid.SelectAll();
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            if (TabUcr.IsSelected)
            {
                if (UCRGrid.SelectedItems.Count == 0)
                {
                    System.Windows.MessageBox.Show("Nothing to Copy");
                    return;
                }
                UCRGrid.Focus();
                UCRGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                ApplicationCommands.Copy.Execute(null, UCRGrid);
                string ResultUcr = (string)System.Windows.Clipboard.GetData(System.Windows.DataFormats.CommaSeparatedValue);
            }
            else if (TabAneStd.IsSelected)
            {
                if (ANEStdDataGrid.SelectedItems.Count == 0)
                {
                    System.Windows.MessageBox.Show("Nothing to Copy");
                    return;
                }
                ANEStdDataGrid.Focus();
                ANEStdDataGrid.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
                ApplicationCommands.Copy.Execute(null, ANEStdDataGrid);
                string ResultUcr = (string)System.Windows.Clipboard.GetData(System.Windows.DataFormats.CommaSeparatedValue);
            }
            else    // Anesthesia Base units
            {
                if (ANEBaseGrid1.SelectedItems.Count == 0)
                {
                    System.Windows.MessageBox.Show("Nothing to Copy");
                    return;
                }
                ANEBaseGrid1.Focus();
                ANEBaseGrid1.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
                ApplicationCommands.Copy.Execute(null, ANEBaseGrid1);
                string ResultUcr = (string)System.Windows.Clipboard.GetData(System.Windows.DataFormats.CommaSeparatedValue);
            }
        }

        private void Print_Click(object sender, RoutedEventArgs e)
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                System.Windows.MessageBox.Show("Printing requires Microsoft Excel");
                return;
            }
            using (OverrideCursor oCursor = new OverrideCursor(System.Windows.Input.Cursors.Wait))
            {
                if (TabUcr.IsSelected)
                {
                    Printing.Print_UCR();
                }

                else if (TabAneStd.IsSelected)
                {
                    Printing.Print_ANESTD();
                }
                else    // Anesthesia Base units
                {
                    Printing.Print_ANEBASE();
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            string sNewFileName, sPath;
            sNewFileName = "";
            sPath = Properties.Settings.Default.DBDefaultPath;


            if (TabUcr.IsSelected)
            {
                if (UCRGrid.SelectedItems.Count == 0)
                {
                    System.Windows.MessageBox.Show("Nothing to Copy");
                    return;
                }

                SaveFileDialog MyExport = new SaveFileDialog();

                MyExport.InitialDirectory = sPath;
                MyExport.FileName = "UCR Export.csv";
                MyExport.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                if (MyExport.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    sNewFileName = MyExport.FileName;
                }
                else    // Canceled out of export
                    return;

                try
                {
                    UCRGrid.Focus();
                    UCRGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                    ApplicationCommands.Copy.Execute(null, UCRGrid);
                    string ResultUcr = (string)System.Windows.Clipboard.GetData(System.Windows.DataFormats.CommaSeparatedValue);
                    System.Windows.Clipboard.Clear();
                    System.IO.StreamWriter file = new System.IO.StreamWriter(sNewFileName);
                    file.WriteLine(ResultUcr);
                    file.Close();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: Could not export UCR to the file: " + sNewFileName + ": " + ex.Message);
                }
            }
            else if (TabAneStd.IsSelected)
            {
                if (ANEStdDataGrid.SelectedItems.Count == 0)
                {
                    System.Windows.MessageBox.Show("Nothing to Copy");
                    return;
                }

                SaveFileDialog MyExport = new SaveFileDialog();

                MyExport.InitialDirectory = sPath;
                MyExport.FileName = "Anesthesia Standard Export.csv";
                MyExport.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                if (MyExport.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    sNewFileName = MyExport.FileName;
                }
                else    // Canceled out of export
                    return;

                try
                {
                    ANEStdDataGrid.Focus();
                    ANEStdDataGrid.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
                    ApplicationCommands.Copy.Execute(null, ANEStdDataGrid);
                    string ResultUcr = (string)System.Windows.Clipboard.GetData(System.Windows.DataFormats.CommaSeparatedValue);
                    System.Windows.Clipboard.Clear();
                    System.IO.StreamWriter file = new System.IO.StreamWriter(sNewFileName);
                    file.WriteLine(ResultUcr);
                    file.Close();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: Could not export Anesthesia Std to the file: " + sNewFileName + ": " + ex.Message);
                }
            }
            else // Ane Base units
            {
                if (ANEBaseGrid1.SelectedItems.Count == 0)
                {
                    System.Windows.MessageBox.Show("Nothing to Copy");
                    return;
                }

                SaveFileDialog MyExport = new SaveFileDialog();

                MyExport.InitialDirectory = sPath;
                MyExport.FileName = "Anesthesia Base Units Export.csv";
                MyExport.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                if (MyExport.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    sNewFileName = MyExport.FileName;
                }
                else    // Canceled out of export
                    return;

                try
                {
                    ANEBaseGrid1.Focus();
                    ANEBaseGrid1.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
                    ApplicationCommands.Copy.Execute(null, ANEBaseGrid1);
                    string ResultUcr = (string)System.Windows.Clipboard.GetData(System.Windows.DataFormats.CommaSeparatedValue);
                    System.Windows.Clipboard.Clear();
                    System.IO.StreamWriter file = new System.IO.StreamWriter(sNewFileName);
                    file.WriteLine(ResultUcr);
                    file.Close();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: Could not export Anesthesia Base Units to the file: " + sNewFileName + ": " + ex.Message);
                }
            }
        }

        private void SelectAll_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            UCRGrid.SelectAll();
        }

        public void SaveColumnSettings()
        {
            Properties.Settings.Default.V25 = bV25th;
            Properties.Settings.Default.V30 = bV30th;
            Properties.Settings.Default.V35 = bV35th;
            Properties.Settings.Default.V40 = bV40th;
            Properties.Settings.Default.V45 = bV45th;
            Properties.Settings.Default.V50 = bV50th;
            Properties.Settings.Default.V55 = bV55th;
            Properties.Settings.Default.V60 = bV60th;
            Properties.Settings.Default.V65 = bV65th;
            Properties.Settings.Default.V70 = bV70th;
            Properties.Settings.Default.V75 = bV75th;
            Properties.Settings.Default.V80 = bV80th;
            Properties.Settings.Default.V85 = bV85th;
            Properties.Settings.Default.V90 = bV90th;
            Properties.Settings.Default.V95 = bV95th;
            Properties.Settings.Default.Save();
        }

        public void GetColumnSettings()
        {
            bV25th = Properties.Settings.Default.V25;
            bV30th = Properties.Settings.Default.V30;
            bV35th = Properties.Settings.Default.V35;
            bV40th = Properties.Settings.Default.V40;
            bV45th = Properties.Settings.Default.V45;
            bV50th = Properties.Settings.Default.V50;
            bV55th = Properties.Settings.Default.V55;
            bV60th = Properties.Settings.Default.V60;
            bV65th = Properties.Settings.Default.V65;
            bV70th = Properties.Settings.Default.V70;
            bV75th = Properties.Settings.Default.V75;
            bV80th = Properties.Settings.Default.V80;
            bV85th = Properties.Settings.Default.V85;
            bV90th = Properties.Settings.Default.V90;
            bV95th = Properties.Settings.Default.V95;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            SaveColumnSettings();
            base.OnClosing(e);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            sImportWorkingPath = Properties.Settings.Default.ImportDefaultPath;
            //
        }

        private void DoFont_Click(object sender, RoutedEventArgs e)
        {
            DialogResult fontResult;
            System.Windows.Forms.FontDialog fontDialog1 = null;
            fontDialog1 = new System.Windows.Forms.FontDialog();
            //fontDialog1.ShowColor = true;

            fontDialog1.Font = Properties.Settings.Default.AppFont;
            //fontDialog1.Color = textBox1.ForeColor;
            fontDialog1.ShowEffects = false;
            fontDialog1.FontMustExist = true;
            fontResult = fontDialog1.ShowDialog();

            if (fontResult == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.AppFont = fontDialog1.Font;
                string newFontFamily = fontDialog1.Font.FontFamily.Name;
                Properties.Settings.Default.AppFontFamily = newFontFamily;
                float newFontSize = fontDialog1.Font.Size;
                Properties.Settings.Default.AppFontSize = newFontSize;
                //System.Drawing.Color newFontColor = fontDialog1.Color;

                SetUCRTabFonts();
                this.InvalidateVisual();
            }
        }

        // Create the OnPropertyChanged method to raise the event 
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        private void UCR_Select_Click(object sender, MouseButtonEventArgs e)
        {

        }

        private void Click_TabChange(object sender, SelectionChangedEventArgs e)
        {
            if (TabUcr.IsSelected)
            {
                if (UCRGrid.Items.Count > 0)
                {
                    EditClearRows.IsEnabled = true;
                    EditSelectAll.IsEnabled = true;
                    EditCopy.IsEnabled = true;
                    EditExport.IsEnabled = true;
                }
                else
                {
                    EditClearRows.IsEnabled = false;
                    EditSelectAll.IsEnabled = false;
                    EditCopy.IsEnabled = false;
                    EditExport.IsEnabled = false;
                }
            }
            else if (TabAneStd.IsSelected)
            {
                if (ANEStdDataGrid.Items.Count > 0)
                {
                    EditClearRows.IsEnabled = true;
                    EditSelectAll.IsEnabled = true;
                    EditCopy.IsEnabled = true;
                    EditExport.IsEnabled = true;
                }
                else
                {
                    EditClearRows.IsEnabled = false;
                    EditSelectAll.IsEnabled = false;
                    EditCopy.IsEnabled = false;
                    EditExport.IsEnabled = false;
                }
            }
            else    // Anesthesia Base units
            {
                if (this.ANEBaseGrid1.Items.Count > 0)
                {
                    EditClearRows.IsEnabled = true;
                    EditSelectAll.IsEnabled = true;
                    EditCopy.IsEnabled = true;
                    EditExport.IsEnabled = true;
                }
                else
                {
                    EditClearRows.IsEnabled = false;
                    EditSelectAll.IsEnabled = false;
                    EditCopy.IsEnabled = false;
                    EditExport.IsEnabled = false;
                }
            }

        }

        private void ANEBaseGrid1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void myPrint(System.Windows.Controls.DataGrid Grid2print, System.Data.DataTable Table4Columns)
        {
            //System.Data.DataTable tempTable = Table4Columns;
            var printTable = new Table();
            var rowGroup = new TableRowGroup();
            printTable.RowGroups.Add(rowGroup);
            var header = new TableRow();
            rowGroup.Rows.Add(header);

            foreach (DataRow row in Grid2print.Items)
            {
                var tableRow = new TableRow();
                rowGroup.Rows.Add(tableRow);

                foreach (DataColumn column in Table4Columns.Columns)
                {
                    var value = row[column].ToString();//mayby some formatting is in order
                    var cell = new TableCell(new Paragraph(new Run(value)));
                    tableRow.Cells.Add(cell);
                }
            }

            System.Windows.Controls.PrintDialog printDialog = new System.Windows.Controls.PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                printDialog.PrintDocument(((IDocumentPaginatorSource)printTable).DocumentPaginator, "print");
            }
        }

        private void shortcutKey_Click(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (TabUcr.IsSelected)
            {
                if ((Search.IsEnabled) && (e.Key == Key.S) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                    Search_Click(null, null);
                else if (Search.IsEnabled == false)
                {
                    System.Windows.MessageBox.Show("No database open");
                }
            }
            else if (TabAneStd.IsSelected)
            {
                if ((AneStdSearch.IsEnabled) && (e.Key == Key.S) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                    AneStdSearch_Click(null, null);
                else if (AneStdSearch.IsEnabled == false)
                { 
                    System.Windows.MessageBox.Show("No database open");
                }
            }
            else
            {
                if ((AneBaseSearch.IsEnabled) && (e.Key == Key.S) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                    AneBaseSearch_Click(null, null);
                else if (AneBaseSearch.IsEnabled == false)
                { 
                    System.Windows.MessageBox.Show("No database open");
                }
            }
        }

        private void SetUCRTabFonts()
        {
            this.FontFamily = new System.Windows.Media.FontFamily(Properties.Settings.Default.AppFontFamily);
            this.FontSize = Properties.Settings.Default.AppFontSize;
            ZipcodeIn.FontFamily = this.FontFamily;
            ZipcodeIn.FontSize = this.FontSize;
            CptcodeIn.FontFamily = this.FontFamily;
            CptcodeIn.FontSize = this.FontSize;
            AneStdZipcodeIn.FontFamily = this.FontFamily;
            AneStdZipcodeIn.FontSize = this.FontSize;
            AnestdCptcodeIn.FontFamily = this.FontFamily;
            AnestdCptcodeIn.FontSize = this.FontSize;
            AneStdMinutes.FontFamily = this.FontFamily;
            AneStdMinutes.FontSize = this.FontSize;
            AneStdPSM.FontFamily = this.FontFamily;
            AneStdPSM.FontSize = this.FontSize;
            AneBaseZipcodeIn.FontFamily = this.FontFamily;
            AneBaseZipcodeIn.FontSize = this.FontSize;
            AneBaseCptcodeIn.FontFamily = this.FontFamily;
            AneBaseCptcodeIn.FontSize = this.FontSize;
            AneBaseMinutes.FontFamily = this.FontFamily;
            AneBaseMinutes.FontSize = this.FontSize;
            AneBasePSM.FontFamily = this.FontFamily;
            AneBasePSM.FontSize = this.FontSize;
            menu1.FontFamily = this.FontFamily;
            menu1.FontSize = this.FontSize;
            ModifierIn.FontFamily = this.FontFamily;
            ModifierIn.FontSize = this.FontSize;
        }

        public void Doc_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(sDocFile);
        }

    }


    public class OverrideCursor : IDisposable
    {

        public OverrideCursor(System.Windows.Input.Cursor changeToCursor)
        {
            Mouse.OverrideCursor = changeToCursor;
        }

        #region IDisposable Members

        public void Dispose()
        {
            Mouse.OverrideCursor = null;
        }

        #endregion
    }

    public class FontDetails : INotifyPropertyChanged
    {
        //[NonSerialized]
        private double _fontSize = 12.0;
        //[DataMember]
        public double fontSize
        {
            get
            {
                return _fontSize;
            }
            set
            {
                _fontSize = value;
                RaisePropertyChanged();
            }
        }
        // [DataMember]
        public string FontFamilyName { get; set; }

        //[NonSerialized]
        private System.Windows.Media.FontFamily _fontFamily;
        public System.Windows.Media.FontFamily fontFamily
        {
            get
            {
                return _fontFamily;
            }
            set
            {
                _fontFamily = value;
                FontFamilyName = _fontFamily.Source;
                RaisePropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void RaisePropertyChanged(String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        public FontDetails()
        {
            if (string.IsNullOrEmpty(FontFamilyName))
            {
                FontFamilyName = "Calibri";
                fontSize = 12.0;
            }
        }
    }

    public class UIPrinter
    {
        #region Properties

        public Int32 VerticalOffset { get; set; }
        public Int32 HorizontalOffset { get; set; }
        public String Title { get; set; }
        public UIElement Content { get; set; }

        #endregion

        #region Initialization

        public void TimelinePrinter()
        {
            HorizontalOffset = 20;
            VerticalOffset = 20;
            Title = "Print ";
        }

        #endregion

        #region Methods

        public Int32 Print()
        {
            var dlg = new System.Windows.Controls.PrintDialog();
            if (dlg.ShowDialog() == true)
            {
                //---FIRST PAGE---//
                // Size the Grid.
                Content.Measure(new System.Windows.Size(Double.PositiveInfinity,
                                         Double.PositiveInfinity));

                System.Windows.Size sizeGrid = Content.DesiredSize;

                //check the width
                if (sizeGrid.Width > dlg.PrintableAreaWidth)
                {
                    //MessageBoxResult result = System.Windows.MessageBox.Show(Properties.Resources.s_EN_Question_PrintWidth, "Print", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    MessageBoxResult result = System.Windows.MessageBox.Show("Exceeds print area. Try less columns.", "Print", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.No)
                        //throw new PrintAborted(Properties.Resources.s_EN_Info_PrintingAborted);
                        return (0);
                }

                // Position of the grid 
                var ptGrid = new System.Windows.Point(HorizontalOffset, VerticalOffset);

                // Layout of the grid
                Content.Arrange(new Rect(ptGrid, sizeGrid));

                //print
                dlg.PrintVisual(Content, Title);

                //---MULTIPLE PAGES---//
                double diff;
                int i = 1;
                while ((diff = sizeGrid.Height - (dlg.PrintableAreaHeight - VerticalOffset * i) * i) > 0)
                {
                    //Position of the grid 
                    var ptSecondGrid = new System.Windows.Point(HorizontalOffset, -sizeGrid.Height + diff + VerticalOffset);

                    // Layout of the grid
                    Content.Arrange(new Rect(ptSecondGrid, sizeGrid));

                    //print
                    int k = i + 1;
                    dlg.PrintVisual(Content, Title + " (Page " + k + ")");

                    i++;
                }

                return i;
            }

            //throw new PrintAborted(Properties.Resources.s_EN_Info_PrintingAborted);
            return (0);
        }

        #endregion
    }


}
