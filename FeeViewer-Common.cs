//using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using System.IO;
//using System.Text;
//using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows;
using System;


namespace FeeViewerPro
{
    class Common
    {
        public static bool ValidateDB(string connectionString)
        {
            //string sSQLZip = "SELECT distinct GeoZip FROM Zip where ZipLow <='006' and '006'<=ziphigh;";
            string sSQLZip = "SELECT distinct DataType FROM Control;";
            string GeoZip = "";
            OleDbConnection conn = null;

            try
            {
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(sSQLZip, conn);
                conn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    GeoZip = (string)reader[0];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(GeoZip) == true)
                {
                    conn.Close();
                    MessageBox.Show("Not a valid database!");
                    return false;
                }
                else
                {
                    conn.Close();
                    return true;
                }
            }
            catch (Exception Ex)
            {
                conn.Close();
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }

        }

        public static bool ValidateCPT(string CPT, string Status, bool bDisplayError)
        {
            string SQL, cptReturn, sStatus;
            var connectionString = Properties.Settings.Default.DbConnectionString;
            OleDbConnection conn = null;

            SQL = "Select cpt, flag from CPT where cpt='" + CPT + "';";
            cptReturn = "";
            try
            {
                sStatus = "";
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(SQL, conn);
                conn.Open();
                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    cptReturn = (string)reader[0];
                    sStatus = (string)reader[1];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(cptReturn) == true)
                {
                    if (bDisplayError)
                    {
                        MessageBox.Show("Invalid CPT!");
                    }
                    conn.Close();
                        return false;
                }
                else
                {
                    conn.Close();
                    Status = sStatus;
                    return true;
                }
            }
            catch (Exception Ex)
            {
                conn.Close();
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }
        }

        public static bool ValidateDRG(string CPT, string Status, bool bDisplayError)
        {
            string SQL, cptReturn, sStatus;
            var connectionString = Properties.Settings.Default.DbConnectionString;
            OleDbConnection conn = null;

            SQL = "Select cpt, flag from CPT where cpt='" + CPT + "';";
            cptReturn = "";
            try
            {
                sStatus = "";
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(SQL, conn);
                conn.Open();
                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    cptReturn = (string)reader[0];
                    sStatus = (string)reader[1];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(cptReturn) == true)
                {
                    if (bDisplayError)
                    {
                        MessageBox.Show("Invalid DRG!");
                    }
                    conn.Close();
                    return false;
                }
                else
                {
                    conn.Close();
                    Status = sStatus;
                    return true;
                }
            }
            catch (Exception Ex)
            {
                conn.Close();
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }
        }

        public static bool ValidateAneCPT(string CPT, string Status, bool bDisplayError)
        {
            string SQL, cptReturn, sStatus;
            var connectionString = Properties.Settings.Default.DbConnectionString;
            OleDbConnection conn = null;

            SQL = "Select cpt, flag from ACPT where cpt='" + CPT + "';";
            cptReturn = "";
            try
            {
                sStatus = "";
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(SQL, conn);
                conn.Open();
                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    cptReturn = (string)reader[0];
                    sStatus = (string)reader[1];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(cptReturn) == true)
                {
                    if (bDisplayError)
                    {
                        MessageBox.Show("Invalid ANE CPT!");
                    }
                    conn.Close();
                    return false;
                }
                else
                {
                    conn.Close();
                    Status = sStatus;
                    return true;
                }
            }
            catch (Exception Ex)
            {
                conn.Close();
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }
        }

        public static bool ValidateModifierCPT(string Cptstart, string Cptend, string Modifier)
        {
            string SQL, modReturn; //SpecCode, 
            var connectionString = Properties.Settings.Default.DbConnectionString;
            OleDbConnection conn = null;

            SQL = "Select cpt from FCT where cpt>='" + Cptstart + "' AND cpt<='" + Cptend + "' AND SpecCode='" + Modifier + "';";
            modReturn = "";
            try
            {
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(SQL, conn);
                conn.Open();
                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    modReturn = (string)reader[0];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(modReturn) == true)
                {
                    conn.Close();
                    MessageBox.Show("Invalid Modifier (modifiers only apply to some codes)!\n\tTo see valid modifers for the code, use the Search All Modifiers\n\nPossible modifiers are: \n\tCPT/HCPCS - blank or 26\n\tHCPCS - blank, NU, RR, UE");
                    return false;
                }
                else
                {
                    conn.Close();
                    return true;
                }
            }
            catch (Exception Ex)
            {
                conn.Close();
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }
        }

        //public static bool ValidateZip(int testzip, ref string GeoZip)
        public static bool ValidateZip(string testzip, ref string GeoZip)
        {
            var connectionString = Properties.Settings.Default.DbConnectionString;
            //string zippart = testzip.ToString();
            string zippart = testzip;
            zippart = zippart.Substring(0, 3);
            string sSQLZip = "SELECT distinct GeoZip FROM Zip where ZipLow <='" + zippart + "' and '" + zippart + "'<=ziphigh;";
            OleDbConnection conn = null;

            try
            {
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(sSQLZip, conn);
                conn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    GeoZip = (string)reader[0];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(GeoZip) == true)
                {
                    MessageBox.Show("Invalid Zipcode!");
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }

        }

        public static bool ValidateAZip(string testzip, ref string GeoZip)
        {
            var connectionString = Properties.Settings.Default.DbConnectionString;
            //string zippart = testzip.ToString();
            string zippart = testzip;
            zippart = zippart.Substring(0, 3);
            string sSQLZip = "SELECT distinct GeoZip FROM AZip where ZipLow <='" + zippart + "' and '" + zippart + "'<=ziphigh;";
            OleDbConnection conn = null;

            try
            {
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(sSQLZip, conn);
                conn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    GeoZip = (string)reader[0];
                }

                // Call Close when done reading.
                reader.Close();
                if (string.IsNullOrEmpty(GeoZip) == true)
                {
                    MessageBox.Show("Invalid Zipcode!");
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return false;
            }

        }

        public static string CPTCRXStatus(string CPT)
        {
            string SQL, SQL2, cptReturn, sStatus;
            var connectionString = Properties.Settings.Default.DbConnectionString;
            OleDbConnection conn = null;
            OleDbConnection conn2 = null;

            SQL = "Select distinct cpt, status from ACRX where cpt='" + CPT + "';";
            cptReturn = "";
            sStatus = "";
            try
            {
                conn = new OleDbConnection(connectionString);
                OleDbCommand myAccessCommand = new OleDbCommand(SQL, conn);
                conn.Open();
                OleDbDataReader reader = myAccessCommand.ExecuteReader();        //conn.ExecuteReader();

                // Call Read before accessing data.
                while (reader.Read())
                {
                    cptReturn = (string)reader[0];
                    sStatus = (string)reader[1];
                    AneStd.returnAne = cptReturn;
                    AneBase.returnAne = cptReturn;
                }

                // Call Close when done reading.
                reader.Close();
                conn.Close();
            }
            catch (Exception Ex)
            {
                conn.Close();
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return sStatus;
            }

            try
            {
                if (string.IsNullOrEmpty(cptReturn) == true)
                {
                    System.Windows.MessageBox.Show("Invalid CPT!");
                    return sStatus;
                }
                else
                {
                    if (sStatus == "IC")     // check if it isn't a singleton and go get a user choice of which ANE code to use
                    {

                        CPT2ANEXWALK WinAneX = new CPT2ANEXWALK(cptReturn);
                        WinAneX.ShowDialog();
                        if (WinAneX.DialogResult == true)
                        {
                            WinAneX.Close();
                        }
                    }
                    else    // get the singleton ANE for the specified CPT - It may return the CPT depending on value entered (already an ANE) or status (BR or NA)
                    {
                        SQL2 = "Select ANE from ACRX where cpt='" + CPT + "';";
                        conn2 = new OleDbConnection(connectionString);
                        OleDbCommand myAccessCommand2 = new OleDbCommand(SQL2, conn2);
                        conn2.Open();
                        OleDbDataReader reader2 = myAccessCommand2.ExecuteReader();        //conn.ExecuteReader();
                        while (reader2.Read())
                        {
                            cptReturn = (string)reader2[0];
                            AneStd.returnAne = cptReturn;
                        }
                        reader2.Close();
                        conn2.Close();

                    }
                    return sStatus;
                }

            }
            catch (Exception Ex)
            {
                conn2.Close();
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                return sStatus;
            }
        }
    }

}
