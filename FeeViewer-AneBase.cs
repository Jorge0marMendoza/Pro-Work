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
using System.ComponentModel;

namespace FeeViewerPro
{
    class AneBase
    {
        public static string returnAne { get; set; }

        public static void AneBaseSearch(string ZipcodeIn, string CptcodeIn, string MinutesIn, string PSMIn)
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
            returnAne = CptcodeIn;

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
                System.Windows.MessageBox.Show("Enter a valid 4 character code (single code only)");
                return;
            }

            if (Common.ValidateAneCPT(CPTCode, CrxStatus, true) == false)
            {
                return;
            }

            returnAne = CPTCode;    // asssume code is a singleton, assign to returnAne
            CrxStatus = Common.CPTCRXStatus(CPTCode);

            RVSID = 778;
            App.myMainWin.ANEBaseData = getAneBaseData(App.myMainWin.AneBaseTable, ZipcodeIn, GEOZIP, CPTCode, RVSID, MinutesIn, PSMIn, returnAne, CrxStatus);
            App.myMainWin.ANEBaseGrid1.ItemsSource = App.myMainWin.ANEBaseData.DefaultView;

        }

        public static System.Data.DataTable getAneBaseData(DataTable AneBaseTableX, string zipcode, string GEOZIP, string cpt, int rvsid, string minutes, string psm,
            string returnAne, string CrxStatus)
        {
            var connectionString = Properties.Settings.Default.DbConnectionString;
            string sSQL, sDescription, sCPTFlag;
            int nRVU;
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
                        MyFillTableANEBase(AneBaseTableX, zipcode, GEOZIP, cpt, sDescription, nRVU, rvsid, minutes, psm, sCPTFlag, sAneReturned, CrxStatus); // need to be able to append to existing table
                    }
                }

                // Call Close when done reading.
                readerANE.Close();
                connANE.Close();
                return AneBaseTableX;
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

        private static void MyFillTableANEBase(DataTable anetbl, string zipcode, string GEOZIP, string cpt, string sDescript, int nRVU, int rvsid, string minutes,
            string psm, string sCPTFlag, string sAne, string sCrxStatus)
        {
            int j, k;
            OleDbConnection connFCT = null;
            var connectionString = Properties.Settings.Default.DbConnectionString;

            string sSQL2 = "SELECT LowBase, HighBase, LowMinute, HighMinute FROM AFCTX where ";
            sSQL2 = sSQL2 + "(cptlow<='" + sAne + "') AND ('" + sAne + "'<=cpthigh) AND GEOZIP='" + GEOZIP + "';";

            try
            {
                connFCT = new OleDbConnection(connectionString);
                OleDbCommand myAccessCmdFCTX = new OleDbCommand(sSQL2, connFCT);
                connFCT.Open();

                OleDbDataReader readerFCTX = myAccessCmdFCTX.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (readerFCTX.Read())
                {
                    MyFillTableValues(anetbl, (IDataRecord)readerFCTX, zipcode, GEOZIP, cpt, sDescript, minutes, psm, sCPTFlag, nRVU, sAne, sCrxStatus);
                }

                // Call Close when done reading.
                readerFCTX.Close();
                connFCT.Close();
                k = anetbl.Rows.Count;
                if (k < 10)
                {
                    DataRow dr = anetbl.NewRow();
                    for (j = k; j < 10; j++)
                    {
                        dr = anetbl.NewRow();
                        dr["Zipcode"] = null;
                        anetbl.Rows.Add(dr);
                    }
                }
            }
            catch (Exception Ex)
            {
                connFCT.Close();
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
            }
            //return dt;
        }

        private static void MyFillTableValues(DataTable AneBaseTBL, IDataRecord record, string zipCode, string GEOZIP, string cpt, string Description, string minutes,
            string psm, string sCPTFlag, int RVU, string sAne, string sCrxStatus)
        {
            DataRow IndvRow = AneBaseTBL.NewRow();
            float dHBHM, dHBLM, dLBHM, dLBLM;
            string sLowBase, sHighBase, sLowMinute, sHighMinute;
            double minutesIn, dBaseHigh, dBaseLow, dMinHigh, dMinLow;
            float minutesCalulated,RoundTenths;
            float psmfloat, basefloat;
            int CurrentNumRows, k;
            int nMinutes, nRemainder, nFinalMinutes;

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

            minutesCalulated = ((float)nFinalMinutes + RoundTenths) * 15.0f;

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
            basefloat = (float)RVU * x1 + psmfloat;

            CurrentNumRows = AneBaseTBL.Rows.Count;
            for (k = CurrentNumRows - 1; k >= 0; k--)
            {
                if (AneBaseTBL.Rows[k].IsNull("Zipcode"))
                    AneBaseTBL.Rows[k].Delete();
                else
                    break;
            }

            try
            {
                // this will handle null values in the database
                if (record.IsDBNull(0))     //Low Base
                    sLowBase = "0";
                else
                    sLowBase = Convert.ToString(record[0]);
                if (double.TryParse(sLowBase, out dBaseLow) == false)
                {
                    System.Windows.MessageBox.Show("Invalid LowBase value");
                    return;
                }

                if (record.IsDBNull(1))
                    sHighBase = "0";
                else
                    sHighBase = Convert.ToString(record[1]);
                if (double.TryParse(sHighBase, out dBaseHigh) == false)
                {
                    System.Windows.MessageBox.Show("Invalid HighBase value");
                    return;
                }

                if (record.IsDBNull(2))
                    sLowMinute = "0";
                else
                    sLowMinute = Convert.ToString(record[2]);
                if (double.TryParse(sLowMinute, out dMinLow) == false)
                {
                    System.Windows.MessageBox.Show("Invalid LowMinute value");
                    return;
                }

                if (record.IsDBNull(3))
                    sHighMinute = "0";
                else
                    sHighMinute = Convert.ToString(record[3]);
                if (double.TryParse(sHighMinute, out dMinHigh) == false)
                {
                    System.Windows.MessageBox.Show("Invalid HighMinute value");
                    return;
                }

                //dHBHM = ((basefloat * (float)dBaseHigh) + ((float)minutesIn * (float)dMinHigh)) * x2;    //High Base / High Minute
                //dHBLM = ((basefloat * (float)dBaseHigh) + ((float)minutesIn * (float)dMinLow)) * x2;    //High Base / High Minute
                //dLBHM = ((basefloat * (float)dBaseLow) + ((float)minutesIn * (float)dMinHigh)) * x2;    //High Base / High Minute
                //dLBLM = ((basefloat * (float)dBaseLow) + ((float)minutesIn * (float)dMinLow)) * x2;    //High Base / High Minute
                dHBHM = ((basefloat * (float)dBaseHigh) + (minutesCalulated * (float)dMinHigh)) * x2;    //High Base / High Minute
                dHBLM = ((basefloat * (float)dBaseHigh) + (minutesCalulated * (float)dMinLow)) * x2;    //High Base / High Minute
                dLBHM = ((basefloat * (float)dBaseLow) + (minutesCalulated * (float)dMinHigh)) * x2;    //High Base / High Minute
                dLBLM = ((basefloat * (float)dBaseLow) + (minutesCalulated * (float)dMinLow)) * x2;    //High Base / High Minute

                IndvRow["Zipcode"] = zipCode;
                IndvRow["Code"] = cpt;
                IndvRow["Type"] = sCrxStatus;
                IndvRow["ANE"] = sAne;
                IndvRow["Status"] = sCPTFlag;
                IndvRow["Description"] = Description.TrimEnd(' ');
                //IndvRow["Minutes"] = minutesCalulated.ToString();
                IndvRow["Minutes"] = minutes;
                IndvRow["PSM"] = psm;
                IndvRow["High Base High Minute"] = dHBHM.ToString("N");
                IndvRow["High Base Low Minute"] = dHBLM.ToString("N");
                IndvRow["Low Base High Minute"] = dLBHM.ToString("N");
                IndvRow["Low Base Low Minute"] = dLBLM.ToString("N");
                AneBaseTBL.Rows.Add(IndvRow);
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
            }
        }

    }
}
