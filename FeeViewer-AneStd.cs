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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Data.OleDb;

namespace FeeViewerPro
{
    class AneStd
    {
        public static string returnAne
        {
            get;
            set;
        }

        public static void AneStdSearch(string ZipcodeIn, string CptcodeIn, string MinutesIn, string PSMIn)
        {
            /* RVSID is 
             * 778 - ANE
            */
            bool bIsNum;
            int testzip, RVSID;
            bool bValidZip;
            string CPTCode, GEOZIP, CrxStatus;

            GEOZIP = "";
            CrxStatus = "";
            returnAne = CptcodeIn;    // asssume code is a singleton, assign to returnAne

            if (((bIsNum = int.TryParse(ZipcodeIn, out testzip)) != true) || (ZipcodeIn.Length != 5))
            {
                System.Windows.MessageBox.Show("Please enter a valid Zipcode!");
                return;
            }
            else
            {
                //bValidZip = Common.ValidateZip(testzip, ref GEOZIP);
                bValidZip = Common.ValidateAZip(ZipcodeIn, ref GEOZIP);
                if (bValidZip == false)
                    return;
            }
            CPTCode = CptcodeIn;
            if (CPTCode.Length != 5)
            {
                System.Windows.MessageBox.Show("Enter a valid CPT code (single code only");
                return;
            }

            if (Common.ValidateAneCPT(CPTCode, CrxStatus, true) == false)
            {
                return;
            }

            CrxStatus = Common.CPTCRXStatus(CPTCode);

            RVSID = 778;
            App.myMainWin.ANEStdData = getAneStddata(App.myMainWin.AneStdTable, ZipcodeIn, GEOZIP, CPTCode, RVSID, MinutesIn, PSMIn, returnAne, CrxStatus);
            App.myMainWin.ANEStdDataGrid.ItemsSource = App.myMainWin.ANEStdData.DefaultView;

        }

        public static DataTable getAneStddata(DataTable AneStdTable, string zipcode, string GEOZIP, string cpt, int rvsid, string minutes, string psm,
            string returnAne, string CrxStatus)
        {
            var connectionString = Properties.Settings.Default.DbConnectionString;
            string sSQL, sDescription, sCPTFlag;
            int nRVU, j, k;
            string sAneReturned;

            OleDbConnection connANE = null;

            sSQL = "SELECT ACPT.CPT, ACPT.FLAG, ACPT.description, arvs.rvu FROM ACPT INNER JOIN ARVS ON ACPT.CPT = ARVS.CPT where acpt.cpt='" + returnAne + "' and arvs.rvsid='" + rvsid + "' Order By ACPT.CPT";

            try
            {
                connANE = new OleDbConnection(connectionString);

                OleDbCommand myAccessCmdANE = new OleDbCommand(sSQL, connANE);
                connANE.Open();

                OleDbDataReader readerANE = myAccessCmdANE.ExecuteReader();

                // Call Read before accessing data.
                sAneReturned = "";
                while (readerANE.Read())
                {
                    sAneReturned = (string)readerANE[0];
                    if (string.IsNullOrEmpty(sAneReturned) == true)
                        System.Windows.MessageBox.Show("Invalid CPT: " + cpt);
                    else
                    {
                        sDescription = "";
                        nRVU = 0;
                        sCPTFlag = "";
                        GetDescANE((IDataRecord)readerANE, ref sDescription, ref nRVU, ref sCPTFlag);
                        MyFillTableANE(AneStdTable, zipcode, GEOZIP, cpt, sDescription, nRVU, rvsid, minutes, psm, sCPTFlag, sAneReturned, CrxStatus); // need to be able to append to existing table
                    }
                }
                k = AneStdTable.Rows.Count;
                if (k < 10)
                {
                    DataRow dr = AneStdTable.NewRow();
                    for (j = k; j < 10; j++)
                    {
                        dr = AneStdTable.NewRow();
                        dr["Zipcode"] = null;
                        AneStdTable.Rows.Add(dr);
                    }
                }

                // Call Close when done reading.
                readerANE.Close();
                connANE.Close();
                return AneStdTable;
            }
            catch (Exception Ex)
            {
                connANE.Close();
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return null;
            }
        }

        private static void GetDescANE(IDataRecord record, ref string sDesc, ref int Rvu, ref string sReturnFlag)
        {
            sReturnFlag = (string)record[1];
            sDesc = (string)record[2];
            Rvu = (int)record[3];
        }

        private static void MyFillTableANE(DataTable anetbl, string zipcode, string GEOZIP, string cpt, string sDescript, int nRVU, int rvsid, string minutes, string psm,
            string sCPTFlag, string sAne, string sCrxStatus)
        {
            string SpecCode;
            OleDbConnection connFCT = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            SpecCode = "00";
            string sSQL2 = "SELECT [25th], [30th], [35th], [40th], [45th], [50th], [55th], [60th], [65th], [70th], [75th], [80th], [85th], [90th], [95th] FROM AFCTS where ";
            sSQL2 = sSQL2 + " SpecCode='" + SpecCode + "' AND (cptlow<='" + sAne + "') AND ('" + sAne + "'<=cpthigh) AND GEOZIP='" + GEOZIP + "';";

            try
            {
                connFCT = new OleDbConnection(connectionString);
                OleDbCommand myAccessCmdFCT = new OleDbCommand(sSQL2, connFCT);
                connFCT.Open();

                OleDbDataReader readerFCT = myAccessCmdFCT.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (readerFCT.Read())
                {
                    MyFillTableValues(anetbl, (IDataRecord)readerFCT, zipcode, GEOZIP, cpt, sDescript, minutes, psm, sCPTFlag, nRVU, sAne, sCrxStatus);
                }

                // Call Close when done reading.
                readerFCT.Close();
                connFCT.Close();
            }
            catch (Exception Ex)
            {
                connFCT.Close();
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
            }
            //return dt;
        }

        private static void MyFillTableValues(DataTable UCRTBL, IDataRecord record, string zipCode, string GEOZIP, string cpt, string Description, string minutes,
            string psm, string sCPTFlag, int RVU, string sAne, string sCrxStatus)
        {
            DataRow IndvRow = UCRTBL.NewRow();
            float d25th, d30th, d35th, d40th, d45th, d50th, d55th, d60th, d65th, d70th, d75th, d80th, d85th, d90th, d95th;
            //int i25th, i30th, i35th, i40th, i45th, i50th, i55th, i60th, i65th, i70th, i75th, i80th, i85th, i90th, i95th;
            string s25th, s30th, s35th, s40th, s45th, s50th, s55th, s60th, s65th, s70th, s75th, s80th, s85th, s90th, s95th;
            double minutesIn;
            float minutesCalulated, RoundTenths;
            float psmfloat, basefloat;
            int nMinutes, nFinalMinutes, nRemainder, nDisplayMinutes, CurrentNumRows, k;

            if (double.TryParse(minutes, out minutesIn) == false)
            {
                System.Windows.MessageBox.Show("Invalid minutes value");
                return;
            }
            nMinutes = Convert.ToInt32(minutes);
            nFinalMinutes = nMinutes / 15;
            RoundTenths = 0.0f;
            nRemainder = (nMinutes - (nFinalMinutes * 15));
            //if (nRemainder > 0)
            //if (nRemainder > 0)
            //if (nRemainder > 7)
            //    nFinalMinutes++;
            if (nRemainder > 0 && nRemainder < 3)
                RoundTenths = 0.1f;
            else if (nRemainder == 3)
                RoundTenths = 0.2F;
            else if (nRemainder > 3 && nRemainder < 6)
                RoundTenths = 0.3f;
            else if (nRemainder == 6)
                RoundTenths = 0.4f;
            else if (nRemainder > 6 && nRemainder < 9)
                RoundTenths = 0.5f;
            else if (nRemainder == 9)
                RoundTenths = 0.6f;
            else if (nRemainder > 9 && nRemainder < 12)
                RoundTenths = 0.7f;
            else if (nRemainder == 12)
                RoundTenths = 0.8f;
            else if (nRemainder > 12 && nRemainder < 15)
                RoundTenths = 0.9f;
            else RoundTenths = 0.0f;

            minutesCalulated = (float)nFinalMinutes + RoundTenths;
            nDisplayMinutes = nFinalMinutes * 15;

            if ((psm == "P1") || (psm == "P2") || (psm == "P6"))
                psmfloat = 0.00F;
            else if (psm == "P3")
                psmfloat = 1.0F;
            else if (psm == "P4")
                psmfloat = 2.0F;
            else
                psmfloat = 3.0F;

            float x1 = 0.01F;     // Non Inpatient offset
            float x2 = 0.001F;
            basefloat = (float)RVU * x1 + psmfloat + minutesCalulated ;

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
                IndvRow["Zipcode"] = zipCode;
                IndvRow["Code"] = cpt;
                IndvRow["Type"] = sCrxStatus;
                IndvRow["ANE"] = sAne;
                IndvRow["Status"] = sCPTFlag;
                IndvRow["Description"] = Description.TrimEnd(' ');
                //IndvRow["Minutes"] = nDisplayMinutes.ToString();
                IndvRow["Minutes"] = minutes;
                IndvRow["PSM"] = psm;

                if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(0))
                        s25th = "0";
                    else
                        s25th = Convert.ToString(record[0]);
                    d25th = basefloat * (float)Convert.ToSingle(s25th) * x2;    //25th
                    IndvRow["25th"] = d25th.ToString("N");
                }
                else
                    IndvRow["25th"] = "";

                if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(1))
                        s30th = "0";
                    else
                        s30th = Convert.ToString(record[1]);
                    d30th = basefloat * (float)Convert.ToSingle(s30th) * x2;    //30th
                    IndvRow["30th"] = d30th.ToString("N");
                }
                else
                    IndvRow["30th"] = "";

                if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(2))
                        s35th = "0";
                    else
                        s35th = Convert.ToString(record[2]);
                    d35th = basefloat * (float)Convert.ToSingle(s35th) * x2;    //35th
                    IndvRow["35th"] = d35th.ToString("N");
                }
                else
                    IndvRow["35th"] = "";

                if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(3))
                        s40th = "0";
                    else
                        s40th = Convert.ToString(record[3]);
                    d40th = basefloat * (float)Convert.ToSingle(s40th) * x2;    //40th
                    IndvRow["40th"] = d40th.ToString("N");
                }
                else
                    IndvRow["40th"] = "";

                if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(4))
                        s45th = "0";
                    else
                        s45th = Convert.ToString(record[4]);
                    d45th = basefloat * (float)Convert.ToSingle(s45th) * x2;    //45th
                    IndvRow["45th"] = d45th.ToString("N");
                }
                else
                    IndvRow["45th"] = "";

                if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(5))
                        s50th = "0";
                    else
                        s50th = Convert.ToString(record[5]);
                    d50th = basefloat * (float)Convert.ToSingle(s50th) * x2;    //50th
                    IndvRow["50th"] = d50th.ToString("N");
                }
                else
                    IndvRow["50th"] = "";

                if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(6))
                        s55th = "0";
                    else
                        s55th = Convert.ToString(record[6]);
                    d55th = basefloat * (float)Convert.ToSingle(s55th) * x2;    //55th
                    IndvRow["55th"] = d55th.ToString("N");
                }
                else
                    IndvRow["55th"] = "";

                if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(7))
                        s60th = "0";
                    else
                        s60th = Convert.ToString(record[7]);
                    d60th = basefloat * (float)Convert.ToSingle(s60th) * x2;    //60th
                    IndvRow["60th"] = d60th.ToString("N");
                }
                else
                    IndvRow["60th"] = "";

                if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(8))
                        s65th = "0";
                    else
                        s65th = Convert.ToString(record[8]);
                    d65th = basefloat * (float)Convert.ToSingle(s65th) * x2;    //65th
                    IndvRow["65th"] = d65th.ToString("N");
                }
                else
                    IndvRow["65th"] = "";

                if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(9))
                        s70th = "0";
                    else
                        s70th = Convert.ToString(record[9]);
                    d70th = basefloat * (float)Convert.ToSingle(s70th) * x2;    //70th
                    IndvRow["70th"] = d70th.ToString("N");
                }
                else
                    IndvRow["70th"] = "";

                if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(10))
                        s75th = "0";
                    else
                        s75th = Convert.ToString(record[10]);
                    d75th = basefloat * (float)Convert.ToSingle(s75th) * x2;    //75th
                    IndvRow["75th"] = d75th.ToString("N");
                }
                else
                    IndvRow["75th"] = "";

                if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(11))
                        s80th = "0";
                    else
                        s80th = Convert.ToString(record[11]);
                    d80th = basefloat * (float)Convert.ToSingle(s80th) * x2;    //80th
                    IndvRow["80th"] = d80th.ToString("N");
                }
                else
                    IndvRow["80th"] = "";

                if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(12))
                        s85th = "0";
                    else
                        s85th = Convert.ToString(record[12]);
                    d85th = basefloat * (float)Convert.ToSingle(s85th) * x2;    //85th
                    IndvRow["85th"] = d85th.ToString("N");
                }
                else
                    IndvRow["85th"] = "";

                if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(13))
                        s90th = "0";
                    else
                        s90th = Convert.ToString(record[13]);
                    d90th = basefloat * (float)Convert.ToSingle(s90th) * x2;    //90th
                    IndvRow["90th"] = d90th.ToString("N");
                }
                else
                    IndvRow["90th"] = "";

                if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                {
                    if (record.IsDBNull(14))
                        s95th = "0";
                    else
                        s95th = Convert.ToString(record[14]);
                    d95th = basefloat * (float)Convert.ToSingle(s95th) * x2;    //95th
                    IndvRow["95th"] = d95th.ToString("N");
                }
                else
                    IndvRow["95th"] = "";

                UCRTBL.Rows.Add(IndvRow);
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
            }
        }

    }
}
