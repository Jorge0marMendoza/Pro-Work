using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using System.IO;

namespace FeeViewerPro
{
    class DBAccess
    {
        //int NumRowsFCT;

        //public static bool getUCRdata(DataTable UcrTable, int zipcode, string cpt, int rvsid, string modifier, string GEOZIP, string SelectType, bool bisRange, string cpt2)
        public static bool getUCRdata(DataTable UcrTable, string zipcode, string cpt, int rvsid, string modifier, string GEOZIP, string SelectType, bool bisRange, string cpt2)
        {
            var connectionString = Properties.Settings.Default.DbConnectionString;
            string sSQL, sDescription, sReturnModifier, sCPTFlag, sRVSFlag;
            int nRVU, NumRowsRVS;
            int NumRowsFCT, NumRows;
            string sCPTReturned, sStatus;
            DataTable CodeData = new DataTable();
            bool bFCT;
            int j, k;

            OleDbConnection connRVS = null;

            NumRowsFCT = 0;
            NumRows = 0;

            if (bisRange == false)
            {
                sSQL = "SELECT CPT.CPT, CPT.FLAG, CPT.description, rvs.rvu, rvs.modifier, rvs.flag as RVSFlag FROM CPT INNER JOIN RVS ON CPT.CPT = RVS.CPT where cpt.cpt='" + cpt + "' and rvs.rvsid='" + rvsid + "'";
            }
            else
            {
                sSQL = "SELECT CPT.CPT, CPT.FLAG, CPT.description, rvs.rvu, rvs.modifier, rvs.flag as RVSFlag FROM CPT INNER JOIN RVS ON CPT.CPT = RVS.CPT where cpt.cpt>='" + cpt + "' and cpt.cpt<='" + cpt2 + "' and rvs.rvsid='" + rvsid + "'";
            }
            if (rvsid == 775)  //HCPCS Modifiers are in the RVS table
            {
                if (App.myMainWin.CheckAllMods.IsChecked == false) // only need to add modifier to query if CheckAllMods is not checked
                {
                    sSQL = sSQL + " AND rvs.Modifier='" + modifier + "'";
                }
            }
            sSQL = sSQL + " Order By CPT.CPT";

            try
            {
                connRVS = new OleDbConnection(connectionString);

                OleDbCommand myAccessCmdRVS = new OleDbCommand(sSQL, connRVS);
                connRVS.Open();

                var da = new OleDbDataAdapter(myAccessCmdRVS);
                NumRowsRVS = da.Fill(CodeData);
            }
            catch (Exception Ex)
            {
                connRVS.Close();
                MessageBox.Show(Ex.Message.ToString(), "Error getting procedure and RVS table data");
                return false;
            }
            if (NumRowsRVS < 1)
            {
                MessageBox.Show("Unable to get the list of codes for the specified procedure code");
                return false;
            }

            // Call Read before accessing data.
            sCPTReturned = "";
            sStatus = "";
            sDescription = "";
            nRVU = 0;
            //NumRowsFCT = 0;
            sReturnModifier = ""; // only HCPCS has a value here blank, NU, RR or UE
            sCPTFlag = "";
            try
            {
                for (int i = 0; i < NumRowsRVS; i++)
                {
                    sCPTReturned = (string)CodeData.Rows[i]["CPT"];
                    sDescription = (string)CodeData.Rows[i]["Description"];
                    nRVU = (int)CodeData.Rows[i]["RVU"];
                    sReturnModifier = (string)CodeData.Rows[i]["Modifier"]; // only HCPCS has a value here blank, NU, RR or UE
                    sRVSFlag = (string)CodeData.Rows[i]["RVSFlag"];
                    sRVSFlag = sRVSFlag.Trim();
                    sCPTFlag = (string)CodeData.Rows[i]["Flag"];
                    sCPTFlag = sCPTFlag.Trim();
                    if (string.IsNullOrEmpty(sCPTFlag)) // CPT flag is empty
                    {
                        sCPTFlag = sRVSFlag;
                    }
                    //else  // Use the CPT flag
                    //{
                    //    sCPTFlag = (string)CodeData.Rows[i]["Flag"];

                    //}
                    sCPTFlag = sCPTFlag.Trim();
                    if (string.IsNullOrEmpty(sCPTReturned) == true) // This is outside the while loop as it falls here when no data is returned.
                    {
                        MessageBox.Show("Invalid modifier for this code");
                        return false;
                    }
                    if (sCPTFlag == "D1")
                    {
                        sStatus = "Deleted - new";
                    }
                    else if (sCPTFlag == "D2")
                    {
                        sStatus = "Deleted - previous";
                    }
                    else if (sCPTFlag == "N")
                    {
                        sStatus = "New";
                    }
                    else if (sCPTFlag == "U")
                    {
                        sStatus = "Unlisted";
                    }
                    else if (sCPTFlag == "BR")
                    {
                        sStatus = "By Report";
                    }
                    else if (sCPTFlag == "NR")
                    {
                        sStatus = "Rel. Value not est.";
                    }
                    else if (sCPTFlag == "NU")
                    {
                        sStatus = "New Unlisted";
                    }
                    else
                    {
                        sStatus = sCPTFlag;
                    }
                    NumRows = 0;
                    bFCT = MyFillTableUCR(UcrTable, zipcode, sCPTReturned, sDescription, nRVU, rvsid, modifier, GEOZIP, SelectType, sReturnModifier, sStatus, ref NumRows); // need to be able to append to existing table
                    if (bFCT == true)
                        NumRowsFCT = NumRowsFCT + NumRows;
                }
                k = UcrTable.Rows.Count;
                if (k < 10)
                {
                    DataRow dr = UcrTable.NewRow();
                    for (j = k; j < 10; j++)
                    {
                        dr = UcrTable.NewRow();
                        dr["Zipcode"] = null;
                        UcrTable.Rows.Add(dr);
                    }
                }
            }
            catch (Exception Ex)
            {
                connRVS.Close();
                MessageBox.Show(Ex.Message.ToString(), "Error processing Code, Description, RVU");
                return false;
            }

            App.myMainWin.UcrSearchCount.Text = NumRowsFCT.ToString();

            // Call Close when done reading.
            connRVS.Close();
            return true;
        }

        private static void GetDescRVU(IDataRecord record, ref string sDesc, ref int Rvu, ref string sReturnModifier, ref string sReturnFlag)
        {
            sReturnFlag = (string)record[1];
            sDesc = (string)record[2];
            Rvu = (int)record[3];
            sReturnModifier = (string)record[4];
        }

        //private static bool MyFillTableUCR(DataTable ucrtbl, int zipcode, string cpt, string sDescript, int nRVU, int rvsid, string modifier, string GEOZIP,
        private static bool MyFillTableUCR(DataTable ucrtbl, string zipcode, string cpt, string sDescript, int nRVU, int rvsid, string modifier, string GEOZIP,
            string SelectType, string sReturnModifier, string sCPTFlag, ref int NumRows)
        {
            string SpecCode, ActualModifier;
            //int NumRows;
            OleDbConnection connFCT = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;
            DataTable FCTData = new DataTable();

            string sSQL2 = "SELECT [25th], [30th], [35th], [40th], [45th], [50th], [55th], [60th], [65th], [70th], [75th], [80th], [85th], [90th], [95th], SpecCode FROM FCT where ";
            if (rvsid == 777)
            {
                if (modifier == "26") // Medical - mod 26
                    SpecCode = " = '03";
                else if (modifier == "*")
                    SpecCode = " LIKE '0[03]";
                else
                    SpecCode = " = '00";
            }
            else if (rvsid == 776) // Dental
                SpecCode = " = '00";
            else if (rvsid == 775) // HCPCS
                SpecCode = " = '00";
            else if (rvsid == 774) // OutPatient
                SpecCode = " = '10";

            // these next 2 will later be hardcoded in the data
            else if (rvsid == 773) // IPP
                SpecCode = " = '12";
            else if (rvsid == 772) // IPD
                SpecCode = " = '11";
            else                    //
                SpecCode = " = ' = '00";
            sSQL2 = sSQL2 + " SpecCode" + SpecCode + "' AND (cptlow<='" + cpt + "') AND ('" + cpt + "'<=cpthigh) AND GEOZIP='" + GEOZIP + "';";

            try
            {
                connFCT = new OleDbConnection(connectionString);
                OleDbCommand myAccessCmdFCT = new OleDbCommand(sSQL2, connFCT);
                connFCT.Open();

                var da = new OleDbDataAdapter(myAccessCmdFCT);
                NumRows = da.Fill(FCTData);
                //if (NumRows > 1)
                //{
                //    MessageBox.Show("more than one");
                //}
                connFCT.Close();
            }
            catch (Exception Ex)
            {
                connFCT.Close();
                MessageBox.Show(Ex.Message.ToString(), "Error reading FCT data");
                return false;
            }
            try
            {
                if (NumRows < 1)
                {
                    return false;
                }
                for (int i = 0; i < NumRows; i++)
                {
                    if (rvsid == 777)
                        ActualModifier = (string)FCTData.Rows[i]["SpecCode"]; //Medical modifiers are in the FCT table
                    else if (rvsid == 775)
                        ActualModifier = sReturnModifier;
                    else
                        ActualModifier = modifier;
                    MyFillTableValues(ucrtbl, i, FCTData, zipcode, cpt, sDescript, modifier, nRVU, SelectType, ActualModifier, sCPTFlag, rvsid);
                }

                connFCT.Close();
                return true;
            }
            catch (Exception Ex)
            {
                connFCT.Close();
                MessageBox.Show(Ex.Message.ToString(), "Error loading FCT data");
                return false;
            }
            //return dt;
        }

        //private static void MyFillTableValues(DataTable UCRTBL, int nIndex, DataTable FCT, int zipCode, string cpt, string Description, string Modifier, int RVU,
        private static void MyFillTableValues(DataTable UCRTBL, int nIndex, DataTable FCT, string zipCode, string cpt, string Description, string Modifier, int RVU,
            string SelectType, string ActualModifier, string sCPTFlag, int rvsid)
        {
            DataRow IndvRow = UCRTBL.NewRow();
            float d25th, d30th, d35th, d40th, d45th, d50th, d55th, d60th, d65th, d70th, d75th, d80th, d85th, d90th, d95th;
            int i25th, i30th, i35th, i40th, i45th, i50th, i55th, i60th, i65th, i70th, i75th, i80th, i85th, i90th, i95th;

            float x = 0.00001F;     // Non Inpatient offset
            int CurrentNumRows, k;

            CurrentNumRows = UCRTBL.Rows.Count;
            for (k = CurrentNumRows - 1; k >= 0; k--)
            {
                if (UCRTBL.Rows[k].IsNull("Zipcode"))
                    UCRTBL.Rows[k].Delete();
                else
                    break;
            }

            try
            {
                if ((SelectType == "IPP") || (SelectType == "IPD")) //Inpatient offset
                    x = 0.001F;

                IndvRow["Zipcode"] = zipCode;
                IndvRow["Code"] = cpt;
                IndvRow["Status"] = sCPTFlag;
                IndvRow["Description"] = Description.TrimEnd(' ');
                if (rvsid == 777)
                {
                    if (ActualModifier == "03")
                        IndvRow["Modifier"] = "26";
                    else
                        IndvRow["Modifier"] = "";
                }
                else
                    IndvRow["Modifier"] = ActualModifier;

                IndvRow["Type"] = SelectType;

                // this will handle null values in the database
                if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                {
                    i25th = (int)FCT.Rows[nIndex]["25th"];
                    d25th = (float)RVU * (float)i25th * x;    //25th
                    IndvRow["25th"] = d25th.ToString("N");
                }
                else
                    IndvRow["25th"] = "";


                if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                {
                    i30th = (int)FCT.Rows[nIndex]["30th"];
                    d30th = (float)RVU * (float)i30th * x;    //30th
                    IndvRow["30th"] = d30th.ToString("N");
                }
                else
                    IndvRow["30th"] = "";


                if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                {
                    i35th = (int)FCT.Rows[nIndex]["35th"];
                    d35th = (float)RVU * (float)i35th * x;    //35th
                    IndvRow["35th"] = d35th.ToString("N");
                }
                else
                    IndvRow["35th"] = "";


                if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                {
                    i40th = (int)FCT.Rows[nIndex]["40th"];
                    d40th = (float)RVU * (float)i40th * x;    //40th
                    IndvRow["40th"] = d40th.ToString("N");
                }
                else
                    IndvRow["40th"] = "";


                if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                {
                    i45th = (int)FCT.Rows[nIndex]["45th"];
                    d45th = (float)RVU * (float)i45th * x;    //45th
                    IndvRow["45th"] = d45th.ToString("N");
                }
                else
                    IndvRow["45th"] = "";


                if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                {
                    i50th = (int)FCT.Rows[nIndex]["50th"];
                    d50th = (float)RVU * (float)i50th * x;    //50th
                    IndvRow["50th"] = d50th.ToString("N");
                }
                else
                    IndvRow["50th"] = "";


                if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                {
                    i55th = (int)FCT.Rows[nIndex]["55th"];
                    d55th = (float)RVU * (float)i55th * x;    //55th
                    IndvRow["55th"] = d55th.ToString("N");
                }
                else
                    IndvRow["55th"] = "";


                if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                {
                    i60th = (int)FCT.Rows[nIndex]["60th"];
                    d60th = (float)RVU * (float)i60th * x;    //60th
                    IndvRow["60th"] = d60th.ToString("N");
                }
                else
                    IndvRow["60th"] = "";


                if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                {
                    i65th = (int)FCT.Rows[nIndex]["65th"];
                    d65th = (float)RVU * (float)i65th * x;    //65th
                    IndvRow["65th"] = d65th.ToString("N");
                }
                else
                    IndvRow["65th"] = "";


                if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                {
                    i70th = (int)FCT.Rows[nIndex]["70th"];
                    d70th = (float)RVU * (float)i70th * x;    //70th
                    IndvRow["70th"] = d70th.ToString("N");
                }
                else
                    IndvRow["70th"] = "";


                if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                {
                    i75th = (int)FCT.Rows[nIndex]["75th"];
                    d75th = (float)RVU * (float)i75th * x;    //75th
                    IndvRow["75th"] = d75th.ToString("N");
                }
                else
                    IndvRow["75th"] = "";


                if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                {
                    i80th = (int)FCT.Rows[nIndex]["80th"];
                    d80th = (float)RVU * (float)i80th * x;    //80th
                    IndvRow["80th"] = d80th.ToString("N");
                }
                else
                    IndvRow["80th"] = "";


                if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                {
                    i85th = (int)FCT.Rows[nIndex]["85th"];
                    d85th = (float)RVU * (float)i85th * x;    //85th
                    IndvRow["85th"] = d85th.ToString("N");
                }
                else
                    IndvRow["85th"] = "";


                if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                {
                    i90th = (int)FCT.Rows[nIndex]["90th"];
                    d90th = (float)RVU * (float)i90th * x;    //90th
                    IndvRow["90th"] = d90th.ToString("N");
                }
                else
                    IndvRow["90th"] = "";


                if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                {
                    i95th = (int)FCT.Rows[nIndex]["95th"];
                    d95th = (float)RVU * (float)i95th * x;    //95th
                    IndvRow["95th"] = d95th.ToString("N");
                }
                else
                    IndvRow["95th"] = "";

                UCRTBL.Rows.Add(IndvRow);
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Error creating percentage values");
            }
        }

        /**********************************************************************
         * Check which modules have data
         * WhichModule
         *      0 - ALL, 1 - MED, 2 - DNT, 3 - HCP, 4 - OUT, 5 - IPD, 6 - IPP
         *      7 - Anesthesia
         ***********************************************************************/
        public static void CheckForInstalledModules(int WhichModule)
        {
            OleDbConnection connRVS = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            string sResponse;
            try
            {
                connRVS = new OleDbConnection(connectionString);
                connRVS.Open();
                if (WhichModule == 0 || WhichModule == 1)
                {
                    /********************************/
                    /* Medical                      */
                    /********************************/
                    string sSQLMED = "SELECT DISTINCT RVSID FROM RVS where RVSID='777'"; // MED
                    OleDbCommand myAccessCmdRVS = new OleDbCommand(sSQLMED, connRVS);

                    sResponse = (string)myAccessCmdRVS.ExecuteScalar();
                    if (sResponse != "777")
                    {
                        App.myMainWin.IMP_MED.IsEnabled = true;
                        App.myMainWin.ButtonMedical.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_MED.IsEnabled = false;
                        App.myMainWin.ButtonMedical.IsEnabled = true;
                    }
                }

                if (WhichModule == 0 || WhichModule == 2)
                {
                    /********************************/
                    /* Dental                       */
                    /********************************/
                    string sSQLDNT = "SELECT DISTINCT RVSID FROM RVS where RVSID='776'"; // DNT
                    OleDbCommand myAccessCmdDNT = new OleDbCommand(sSQLDNT, connRVS);

                    sResponse = (string)myAccessCmdDNT.ExecuteScalar();
                    if (sResponse != "776")
                    {
                        App.myMainWin.IMP_DNT.IsEnabled = true;
                        App.myMainWin.ButtonDental.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_DNT.IsEnabled = false;
                        App.myMainWin.ButtonDental.IsEnabled = true;
                    }
                }

                if (WhichModule == 0 || WhichModule == 3)
                {
                    /********************************/
                    /* HCPCS                        */
                    /********************************/
                    string sSQLHCP = "SELECT DISTINCT RVSID FROM RVS where RVSID='775'"; // HCP
                    OleDbCommand myAccessCmdHCP = new OleDbCommand(sSQLHCP, connRVS);

                    sResponse = (string)myAccessCmdHCP.ExecuteScalar();
                    if (sResponse != "775")
                    {
                        App.myMainWin.IMP_HCP.IsEnabled = true;
                        App.myMainWin.ButtonHCPCS.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_HCP.IsEnabled = false;
                        App.myMainWin.ButtonHCPCS.IsEnabled = true;
                    }
                }

                if (WhichModule == 0 || WhichModule == 4)
                {
                    /********************************/
                    /* Outpatient                   */
                    /********************************/

                    string sSQLOUT = "SELECT DISTINCT RVSID FROM RVS where RVSID='774'"; // OUT
                    OleDbCommand myAccessCmdOUT = new OleDbCommand(sSQLOUT, connRVS);

                    sResponse = (string)myAccessCmdOUT.ExecuteScalar();
                    if (sResponse != "774")
                    {
                        App.myMainWin.IMP_OUT.IsEnabled = true;
                        App.myMainWin.ButtonOP.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_OUT.IsEnabled = false;
                        App.myMainWin.ButtonOP.IsEnabled = true;
                    }

                }

                if (WhichModule == 0 || WhichModule == 5)
                {
                    /********************************/
                    /* Inpatient by Day             */
                    /********************************/
                    string sSQLIPD = "SELECT DISTINCT RVSID FROM RVS where RVSID='772'"; // IPD
                    OleDbCommand myAccessCmdIPD = new OleDbCommand(sSQLIPD, connRVS);

                    sResponse = (string)myAccessCmdIPD.ExecuteScalar();
                    if (sResponse != "772")
                    {
                        App.myMainWin.IMP_IPD.IsEnabled = true;
                        App.myMainWin.ButtonIPD.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_IPD.IsEnabled = false;
                        //App.myMainWin.bHasDRG = true;
                        App.myMainWin.ButtonIPD.IsEnabled = true;
                    }
                }

                if (WhichModule == 0 || WhichModule == 6)
                {
                    /********************************/
                    /* Inpatient by Patient         */
                    /********************************/
                    string sSQLIPP = "SELECT DISTINCT RVSID FROM RVS where RVSID='773'"; // IPP
                    OleDbCommand myAccessCmdIPP = new OleDbCommand(sSQLIPP, connRVS);

                    sResponse = (string)myAccessCmdIPP.ExecuteScalar();
                    if (sResponse != "773")
                    {
                        App.myMainWin.IMP_IPP.IsEnabled = true;
                        App.myMainWin.ButtonIPP.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_IPP.IsEnabled = false;
                        //App.myMainWin.bHasDRG = true;
                        App.myMainWin.ButtonIPP.IsEnabled = true;
                    }
                }

                if (WhichModule == 0 || WhichModule == 7)
                {
                    /********************************/
                    /* Anesthesia                   */
                    /********************************/
                    string sSQLIPP = "SELECT DISTINCT RVSID FROM ARVS where RVSID='778'"; // ANE
                    OleDbCommand myAccessCmdIPP = new OleDbCommand(sSQLIPP, connRVS);

                    sResponse = (string)myAccessCmdIPP.ExecuteScalar();
                    if (sResponse != "778")
                    {
                        App.myMainWin.IMP_ANE.IsEnabled = true;
                        App.myMainWin.TabAneBase.IsEnabled = false;
                        App.myMainWin.TabAneStd.IsEnabled = false;
                    }
                    else
                    {
                        App.myMainWin.IMP_ANE.IsEnabled = false;
                        App.myMainWin.TabAneBase.IsEnabled = true;
                        App.myMainWin.TabAneStd.IsSelected = true;
                        App.myMainWin.TabAneStd.IsEnabled = true;
                    }
                }
                 //set the default button to the first installed module
                App.myMainWin.TabUcr.IsEnabled = false;
                if (App.myMainWin.ButtonMedical.IsEnabled == true)
                {
                    App.myMainWin.ButtonMedical.IsChecked = true;
                    App.myMainWin.TabUcr.IsEnabled = true;
                }
                else if (App.myMainWin.ButtonDental.IsEnabled == true)
                {
                    App.myMainWin.ButtonDental.IsChecked = true;
                    App.myMainWin.TabUcr.IsEnabled = true;
                }
                else if (App.myMainWin.ButtonHCPCS.IsEnabled == true)
                {
                    App.myMainWin.ButtonHCPCS.IsChecked = true;
                    App.myMainWin.TabUcr.IsEnabled = true;
                }
                else if (App.myMainWin.ButtonOP.IsEnabled == true)
                {
                    App.myMainWin.ButtonOP.IsChecked = true;
                    App.myMainWin.TabUcr.IsEnabled = true;
                }
                else if (App.myMainWin.ButtonIPD.IsEnabled == true)
                {
                    App.myMainWin.ButtonIPD.IsChecked = true;
                    App.myMainWin.TabUcr.IsEnabled = true;
                }
                else if (App.myMainWin.ButtonIPP.IsEnabled == true)
                {
                    App.myMainWin.ButtonIPP.IsChecked = true;
                    App.myMainWin.TabUcr.IsEnabled = true;
                }
                if (App.myMainWin.TabUcr.IsEnabled == true)
                {
                    App.myMainWin.TabUcr.IsSelected = true;
                }


                connRVS.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                connRVS.Close();
            }
        }

        public static void InsertMED(string CPTFile)
        {
            string dataPath;

            dataPath = System.IO.Path.GetDirectoryName(CPTFile);
            if (App.myMainWin.bHasCPT == false)
            {
                InsertCpt(dataPath, "ucrcpt");     //UCRType parameter is ucr for Medical
                App.myMainWin.bHasCPT = true;       // this is to prevent trying to install duplicates with Outpatient
            }
            InsertFct(dataPath, "ucr");
            InsertRvs(dataPath, "ucr");
            InsertZip(dataPath, "ucr");
        }

        public static void InsertDNT(string DNTFile)
        {
            string dataPath;

            dataPath = System.IO.Path.GetDirectoryName(DNTFile);
            InsertCpt(dataPath, "dntcpt");     //UCRType parameter is ucr for Dental
            InsertFct(dataPath, "dnt");
            InsertRvs(dataPath, "dnt");
            InsertZip(dataPath, "dnt");

        }

        public static void InsertHCP(string HCPFile)
        {
            string dataPath;

            dataPath = System.IO.Path.GetDirectoryName(HCPFile);
            InsertCpt(dataPath, "hpxcpt");     //UCRType parameter is ucr for HCPCS
            InsertFct(dataPath, "hpx");
            InsertRvs(dataPath, "hpx");
            InsertZip(dataPath, "hpx");

        }

        public static void InsertIPD(string IPDFile)
        {
            string dataPath;

            dataPath = System.IO.Path.GetDirectoryName(IPDFile);
            if (App.myMainWin.bHasDRG == false)
            {
                InsertCpt(dataPath, "daydrg");     //UCRType parameter is ucr for InPatient Say
                App.myMainWin.bHasDRG = true;
            }
            InsertFct(dataPath, "day");
            InsertRvs(dataPath, "day");
            InsertZip(dataPath, "day");

        }

        public static void InsertIPP(string IPPFile)
        {
            string dataPath;

            dataPath = System.IO.Path.GetDirectoryName(IPPFile);
            if (App.myMainWin.bHasDRG == false)
            {
                InsertCpt(dataPath, "patdrg");     //UCRType parameter is ucr for Inpatient Stay
                App.myMainWin.bHasDRG = true;
            }
            InsertFct(dataPath, "pat");
            InsertRvs(dataPath, "pat");
            InsertZip(dataPath, "pat");

        }

        public static void InsertOPF(string OUTFile)
        {
            string dataPath;

            dataPath = System.IO.Path.GetDirectoryName(OUTFile);
            if (App.myMainWin.bHasCPT == false)
            {
                InsertCpt(dataPath, "opfcpt");     //UCRType parameter is ucr for Outpatient
                App.myMainWin.bHasCPT = true;       // this is to prevent trying to install duplicates with Medical
            }
            InsertFct(dataPath, "opf");
            InsertRvs(dataPath, "opf");
            InsertZip(dataPath, "opf");

        }

        public static void InsertCpt(string CPTPath, string UCRType)
        {
            string FQCPTFile;
            string sSQLCPTStart = "Insert Into CPT values (";
            //string sSQLCPTStartIP = "Insert Into CPT (CPT, Description) values (";
            string sSQLCPTEnd = "');";
            string sSQLInsert;
            string sCPT, SDescription, sFlag;
            int nReturn, lineNo;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            lineNo = 0;
            FQCPTFile = CPTPath + "\\" + UCRType + ".txt";
            try
            {
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQCPTFile))
                {
                    lineNo++;
                    sCPT = sLine.Substring(0, 5);   //for inpatient this is 00DRG, for all others this is CPT/HCPCS
                    SDescription = sLine.Substring(5, 48).Replace("'", "''");
                    if (UCRType != "daydrg" && UCRType != "patdrg")
                    {
                        sFlag = sLine.Substring(53, 3);
                        sSQLInsert = sSQLCPTStart + "'" + sCPT + "','" + SDescription + "','" + sFlag + sSQLCPTEnd;
                    }
                    else    // inpatient has no flag field and a shorter description - fill flag with NULL?
                    {
                        sFlag = "   ";
                        sSQLInsert = sSQLCPTStart + "'" + sCPT + "','" + SDescription + "','" + sFlag + sSQLCPTEnd;
                        //sSQLInsert = sSQLCPTStartIP + "'" + sCPT + "','" + SDescription + "," + sSQLCPTEnd;
                    }
                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                conn.Close();
                string MsgCPT = Ex.Message;
                MsgCPT = MsgCPT + " Filename:" + FQCPTFile + " line#:" + lineNo.ToString();
                MessageBox.Show(MsgCPT, "Load Error");
                return;
            }
        }

        public static void InsertFct(string FCTPath, string UCRType)
        {
            string FQFCTFile;
            string sSQLFCTStart = "Insert Into FCT values (";
            string sSQLFCTEnd = "');";
            string sSQLInsert;
            string GeoZip, SpecCode, CptLow, CptHigh;
            string P25th, P30th, P35th, P40th, P45th, P50th, P55th, P60th, P65th, P70th, P75th, P80th, P85th, P90th, P95th;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQFCTFile = FCTPath + "\\" + UCRType + "xfct.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQFCTFile))
                {
                    GeoZip = sLine.Substring(0, 3);
                    if (UCRType == "day")
                    {
                        SpecCode = "11";
                    }
                    else if (UCRType == "pat")
                    {
                        SpecCode = "12";
                    }
                    else
                        SpecCode = sLine.Substring(3, 2);

                    CptLow = sLine.Substring(5, 5);
                    CptHigh = sLine.Substring(10, 5);
                    P25th = sLine.Substring(15, 7);
                    P30th = sLine.Substring(22, 7);
                    P35th = sLine.Substring(29, 7);
                    P40th = sLine.Substring(36, 7);
                    P45th = sLine.Substring(43, 7);
                    P50th = sLine.Substring(50, 7);
                    P55th = sLine.Substring(57, 7);
                    P60th = sLine.Substring(64, 7);
                    P65th = sLine.Substring(71, 7);
                    P70th = sLine.Substring(78, 7);
                    P75th = sLine.Substring(85, 7);
                    P80th = sLine.Substring(92, 7);
                    P85th = sLine.Substring(99, 7);
                    P90th = sLine.Substring(106, 7);
                    P95th = sLine.Substring(113, 7);
                    sSQLInsert = sSQLFCTStart + "'" + GeoZip + "','" + SpecCode + "','" + CptLow + "','" + CptHigh + "','";
                    sSQLInsert = sSQLInsert + P25th + "','" + P30th + "','" + P35th + "','" + P40th + "','" + P45th + "','" + P50th + "','" + P55th + "','";
                    sSQLInsert = sSQLInsert + P60th + "','" + P65th + "','" + P70th + "','" + P75th + "','" + P80th + "','" + P85th + "','";
                    sSQLInsert = sSQLInsert + P90th + "','" + P95th + sSQLFCTEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }

        }

        public static void InsertRvs(string RVSPath, string UCRType)
        {
            string FQRVSFile;
            string sSQLRVSStart = "Insert Into RVS values (";
            string sSQLRVSEnd = "');";
            string sSQLInsert;
            string RVSID, CPT, RVU, Flag, Modifier;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQRVSFile = RVSPath + "\\" + UCRType + "rvs.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQRVSFile))
                {
                    RVSID = sLine.Substring(0, 3);
                    CPT = sLine.Substring(3, 5);
                    RVU = sLine.Substring(8, 7);
                    Flag = sLine.Substring(15, 2);
                    Modifier = sLine.Substring(17, 2);
                    sSQLInsert = sSQLRVSStart + "'" + RVSID + "','" + CPT + "','" + RVU + "','" + Flag + "','" + Modifier + sSQLRVSEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }
        }

        public static void InsertZip(string ZIPPath, string UCRType)
        {
            string FQZIPFile;
            string sSQLZIPStart = "Insert Into ZIP values (";
            string sSQLZIPEnd = "');";
            string sSQLInsert;
            string ZipLow, ZipHigh, GeoZip, Description, ZipAreas, RvsID;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQZIPFile = ZIPPath + "\\" + UCRType + "zip.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQZIPFile))
                {
                    ZipLow = sLine.Substring(0, 3);
                    ZipHigh = sLine.Substring(3, 3);
                    GeoZip = sLine.Substring(6, 3);
                    Description = sLine.Substring(9, 40);
                    ZipAreas = sLine.Substring(49, 40);
                    RvsID = sLine.Substring(89, 3);
                    sSQLInsert = sSQLZIPStart + "'" + ZipLow + "','" + ZipHigh + "','" + GeoZip + "','" + Description + "','" + ZipAreas + "','" + RvsID + sSQLZIPEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }
        }

        public static void InsertANE(string ANEFile)
        {
            string dataPath;
            dataPath = System.IO.Path.GetDirectoryName(ANEFile);
            InsertACpt(dataPath, "anecpt");     //UCRType parameter is ucr for Medical
            InsertAFct(dataPath, "ane");
            InsertARvs(dataPath, "ane");
            InsertAZip(dataPath, "ane");
            InsertACPT2ANE(dataPath, "ane"); // CPT to ANE xwalk
            InsertAFCTX(dataPath, "ane");    // ANE Base units FCT file
        }

        public static void InsertACpt(string CPTPath, string UCRType)
        {
            string FQCPTFile;
            string sSQLCPTStart = "Insert Into ACPT values (";
            string sSQLCPTEnd = "');";
            string sSQLInsert;
            string sCPT, SDescription, sFlag;
            int nReturn, lineNo;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            lineNo = 0;
            FQCPTFile = CPTPath + "\\" + UCRType + ".txt";
            try
            {
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQCPTFile))
                {
                    lineNo++;
                    sCPT = sLine.Substring(0, 5);   //for inpatient this is 00DRG, for all others this is CPT/HCPCS
                    SDescription = sLine.Substring(5, 48).Replace("'", "''");
                    sFlag = sLine.Substring(53, 3);
                    sSQLInsert = sSQLCPTStart + "'" + sCPT + "','" + SDescription + "','" + sFlag + sSQLCPTEnd;
                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                conn.Close();
                string MsgCPT = Ex.Message;
                MsgCPT = MsgCPT + " Filename:" + FQCPTFile + " line#:" + lineNo.ToString();
                MessageBox.Show(MsgCPT, "Load Error");
                return;
            }
        }

        public static void InsertAFct(string FCTPath, string UCRType)
        {
            string FQFCTFile;
            string sSQLFCTStart = "Insert Into AFCTS values (";
            string sSQLFCTEnd = "');";
            string sSQLInsert;
            string GeoZip, SpecCode, CptLow, CptHigh;
            string P25th, P30th, P35th, P40th, P45th, P50th, P55th, P60th, P65th, P70th, P75th, P80th, P85th, P90th, P95th;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQFCTFile = FCTPath + "\\" + UCRType + "xfcts.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQFCTFile))
                {
                    GeoZip = sLine.Substring(0, 3);
                    SpecCode = sLine.Substring(3, 2);

                    CptLow = sLine.Substring(5, 5);
                    CptHigh = sLine.Substring(10, 5);
                    P25th = sLine.Substring(15, 7);
                    P30th = sLine.Substring(22, 7);
                    P35th = sLine.Substring(29, 7);
                    P40th = sLine.Substring(36, 7);
                    P45th = sLine.Substring(43, 7);
                    P50th = sLine.Substring(50, 7);
                    P55th = sLine.Substring(57, 7);
                    P60th = sLine.Substring(64, 7);
                    P65th = sLine.Substring(71, 7);
                    P70th = sLine.Substring(78, 7);
                    P75th = sLine.Substring(85, 7);
                    P80th = sLine.Substring(92, 7);
                    P85th = sLine.Substring(99, 7);
                    P90th = sLine.Substring(106, 7);
                    P95th = sLine.Substring(113, 7);
                    sSQLInsert = sSQLFCTStart + "'" + GeoZip + "','" + SpecCode + "','" + CptLow + "','" + CptHigh + "','";
                    sSQLInsert = sSQLInsert + P25th + "','" + P30th + "','" + P35th + "','" + P40th + "','" + P45th + "','" + P50th + "','" + P55th + "','";
                    sSQLInsert = sSQLInsert + P60th + "','" + P65th + "','" + P70th + "','" + P75th + "','" + P80th + "','" + P85th + "','";
                    sSQLInsert = sSQLInsert + P90th + "','" + P95th + sSQLFCTEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }

        }

        public static void InsertARvs(string RVSPath, string UCRType)
        {
            string FQRVSFile;
            string sSQLRVSStart = "Insert Into ARVS values (";
            string sSQLRVSEnd = "');";
            string sSQLInsert;
            string RVSID, CPT, RVU, Flag, Modifier;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQRVSFile = RVSPath + "\\" + UCRType + "rvs.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQRVSFile))
                {
                    RVSID = sLine.Substring(0, 3);
                    CPT = sLine.Substring(3, 5);
                    RVU = sLine.Substring(8, 7);
                    Flag = sLine.Substring(15, 2);
                    Modifier = sLine.Substring(17, 2);
                    sSQLInsert = sSQLRVSStart + "'" + RVSID + "','" + CPT + "','" + RVU + "','" + Flag + "','" + Modifier + sSQLRVSEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }
        }

        public static void InsertAZip(string ZIPPath, string UCRType)
        {
            string FQZIPFile;
            string sSQLZIPStart = "Insert Into AZIP values (";
            string sSQLZIPEnd = "');";
            string sSQLInsert;
            string ZipLow, ZipHigh, GeoZip, Description, ZipAreas, RvsID;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQZIPFile = ZIPPath + "\\" + UCRType + "zip.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQZIPFile))
                {
                    ZipLow = sLine.Substring(0, 3);
                    ZipHigh = sLine.Substring(3, 3);
                    GeoZip = sLine.Substring(6, 3);
                    Description = sLine.Substring(9, 40);
                    ZipAreas = sLine.Substring(49, 40);
                    RvsID = sLine.Substring(89, 3);
                    sSQLInsert = sSQLZIPStart + "'" + ZipLow + "','" + ZipHigh + "','" + GeoZip + "','" + Description + "','" + ZipAreas + "','" + RvsID + sSQLZIPEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }
        }

        public static void InsertACPT2ANE(string CRXPath, string UCRType)
        {
            string FQCRXFile;
            string sSQLCRXStart = "Insert Into ACRX values (";
            string sSQLCRXEnd = "');";
            string sSQLInsert;
            string AneProcedure, AneAnesthesia, AneStatus;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQCRXFile = CRXPath + "\\" + UCRType + "CRX.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQCRXFile))
                {
                    AneProcedure = sLine.Substring(0, 5);
                    AneAnesthesia = sLine.Substring(5, 5);
                    AneStatus = sLine.Substring(10, 2);
                    sSQLInsert = sSQLCRXStart + "'" + AneProcedure + "','" + AneAnesthesia + "','" + AneStatus + sSQLCRXEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }
        }

        public static void InsertAFCTX(string AFCTXPath, string UCRType)
        {
            string FQAFCTXFile;
            string sSQLAFCTXStart = "Insert Into AFCTX values (";
            string sSQLAFCTXEnd = "');";
            string sSQLInsert;
            string GeoZip, SpecCode, CptLow, CptHigh;
            string lowBaseCF, highBaseCF, lowMinuteCF, highMinuteCF;
            int nReturn;
            OleDbCommand myAccessCommand;

            OleDbConnection conn = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            try
            {
                FQAFCTXFile = AFCTXPath + "\\" + UCRType + "FCTX.txt";
                conn = new OleDbConnection(connectionString);
                conn.Open();
                foreach (string sLine in File.ReadLines(FQAFCTXFile))
                {
                    GeoZip = sLine.Substring(0, 3);
                    SpecCode = sLine.Substring(3, 2);
                    CptLow = sLine.Substring(5, 5);
                    CptHigh = sLine.Substring(10, 5);
                    lowBaseCF = sLine.Substring(15, 7);
                    highBaseCF = sLine.Substring(22, 7);
                    lowMinuteCF = sLine.Substring(29, 7);
                    highMinuteCF = sLine.Substring(36, 7);
                    sSQLInsert = sSQLAFCTXStart + "'" + GeoZip + "','" + SpecCode + "','" + CptLow + "','" + CptHigh + "','" + lowBaseCF + "','";
                    sSQLInsert = sSQLInsert + highBaseCF + "','" + lowMinuteCF + "','" + highMinuteCF + sSQLAFCTXEnd;

                    myAccessCommand = new OleDbCommand(sSQLInsert, conn);
                    nReturn = myAccessCommand.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                conn.Close();
                return;
            }
        }

    }
}
