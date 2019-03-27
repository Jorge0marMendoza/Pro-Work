using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using System.Data;

namespace FeeViewerPro
{
    class Printing
    {
        public static void Print_UCR()
        {
            try
            {
                StringBuilder str = new StringBuilder();
                str.Append(@"&""Calibri,Bold,&14FeeViewer Pro Print");

                Microsoft.Office.Interop.Excel.Application m_objExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook m_objBook = m_objExcel.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet m_objSheet = m_objBook.Worksheets.get_Item(1);

                //Global settings
                m_objSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                m_objSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                m_objSheet.Cells[1, 1].EntireRow.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                // set header and footer settings
                m_objSheet.PageSetup.CenterHeader = str.ToString();
                m_objSheet.PageSetup.LeftFooter = App.myMainWin.OpenDbName.Text;
                m_objSheet.PageSetup.CenterFooter = "&D";
                m_objSheet.PageSetup.RightFooter = "&P/&N";

                // set body formats
                m_objSheet.PageSetup.PrintTitleRows = "$1:$1";  // repeats header row across pages
                m_objSheet.PageSetup.LeftMargin = 20;
                m_objSheet.PageSetup.RightMargin = 20;
                m_objSheet.PageSetup.TopMargin = 50;
                m_objSheet.PageSetup.BottomMargin = 50;
                m_objSheet.PageSetup.Zoom = false;
                m_objSheet.PageSetup.FitToPagesWide = 1;
                m_objSheet.PageSetup.FitToPagesTall = false;
                m_objSheet.PageSetup.PrintGridlines = true;

                for (int i = 7; i < 21; i++)
                {
                    m_objSheet.Cells[2, i].EntireColumn.NumberFormat = "#,##0.00";
                }

                int nHeader = 1;
                m_objSheet.Cells[1, nHeader++] = "Zipcode";
                m_objSheet.Cells[1, nHeader++] = "Code";
                m_objSheet.Cells[1, nHeader++] = "Status";
                m_objSheet.Cells[1, nHeader++] = "Description";
                m_objSheet.Cells[1, nHeader++] = "Modifier";
                m_objSheet.Cells[1, nHeader++] = "Type";
                //foreach (object ob in this.UCRGrid.Columns.Select(cs => cs.Header).ToList())
                if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "25th";
                    nHeader++;
                }
                if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "30th";
                    nHeader++;
                }
                if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "35th";
                    nHeader++;
                }
                if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "340th";
                    nHeader++;
                }
                if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "45th";
                    nHeader++;
                }
                if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "50th";
                    nHeader++;
                }
                if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "55th";
                    nHeader++;
                }
                if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "60th";
                    nHeader++;
                }
                if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "65th";
                    nHeader++;
                }
                if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "70th";
                    nHeader++;
                }
                if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "75th";
                    nHeader++;
                }
                if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "80th";
                    nHeader++;
                }
                if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "85th";
                    nHeader++;
                }
                if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "90th";
                    nHeader++;
                }
                if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "95th";
                    nHeader++;
                }


                int nRow = 2;
                int nColumn;

                foreach (DataRowView field in App.myMainWin.UCRGrid.ItemsSource)
                {
                    //DataRow DR = UCRGrid.Row;
                    DataRowView DR = field;
                    m_objSheet.Cells[nRow, 1] = DR[0]; //Zipcode
                    m_objSheet.Cells[nRow, 2] = "'" + DR[1]; //Code
                    m_objSheet.Cells[nRow, 3] = DR[2]; //Status
                    m_objSheet.Cells[nRow, 4] = DR[3]; //Description
                    m_objSheet.Cells[nRow, 5] = DR[4]; //Modifier
                    m_objSheet.Cells[nRow, 6] = DR[5]; //Type

                    nColumn = 7;
                    if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[6];
                        nColumn++;
                    }
                    if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[7];
                        nColumn++;
                    }
                    if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[8];
                        nColumn++;
                    }
                    if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[9];
                        nColumn++;
                    }
                    if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[10];
                        nColumn++;
                    }
                    if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[11];
                        nColumn++;
                    }
                    if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[12];
                        nColumn++;
                    }
                    if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[13];
                        nColumn++;
                    }
                    if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[14];
                        nColumn++;
                    }
                    if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[15];
                        nColumn++;
                    }
                    if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[16];
                        nColumn++;
                    }
                    if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[17];
                        nColumn++;
                    }
                    if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[18];
                        nColumn++;
                    }
                    if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[19];
                        nColumn++;
                    }
                    if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[20];
                        nColumn++;
                    }
                    nRow++;
                }
                m_objSheet.Columns["A:U"].AutoFit();
                m_objExcel.Visible = true;
                // Save the Workbook and quit Excel.
                //m_objBook.SaveAs("D:\\Temp\\Book2.xlsx");
                //m_objBook.Close(true, m_objOpt, false);
                //m_objExcel.Quit();
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                //m_objExcel.Quit();
                return;
            }
        }

        public static void Print_ANESTD()
        {
            try
            {
                StringBuilder str = new StringBuilder();
                str.Append(@"&""Calibri,Bold,&14FeeViewer Pro Print");

                Microsoft.Office.Interop.Excel.Application m_objExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook m_objBook = m_objExcel.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet m_objSheet = m_objBook.Worksheets.get_Item(1);

                //Global settings
                m_objSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                m_objSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                m_objSheet.Cells[1, 1].EntireRow.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                // set header and footer settings
                m_objSheet.PageSetup.CenterHeader = str.ToString();
                m_objSheet.PageSetup.LeftFooter = App.myMainWin.OpenDbName.Text;
                m_objSheet.PageSetup.CenterFooter = "&D";
                m_objSheet.PageSetup.RightFooter = "&P/&N";

                // set body formats
                m_objSheet.PageSetup.PrintTitleRows = "$1:$1";  // repeats header row across pages
                m_objSheet.PageSetup.LeftMargin = 20;
                m_objSheet.PageSetup.RightMargin = 20;
                m_objSheet.PageSetup.TopMargin = 50;
                m_objSheet.PageSetup.BottomMargin = 50;
                m_objSheet.PageSetup.Zoom = false;
                m_objSheet.PageSetup.FitToPagesWide = 1;
                m_objSheet.PageSetup.FitToPagesTall = false;
                m_objSheet.PageSetup.PrintGridlines = true;

                for (int i = 9; i < 23; i++)
                    m_objSheet.Cells[2, i].EntireColumn.NumberFormat = "#,##0.00";

                int nHeader = 1;
                m_objSheet.Cells[1, nHeader++] = "Zipcode";
                m_objSheet.Cells[1, nHeader++] = "CPT";
                m_objSheet.Cells[1, nHeader++] = "Type";
                m_objSheet.Cells[1, nHeader++] = "ANE";
                m_objSheet.Cells[1, nHeader++] = "Status";
                m_objSheet.Cells[1, nHeader++] = "Description";
                m_objSheet.Cells[1, nHeader++] = "Minutes";
                m_objSheet.Cells[1, nHeader++] = "PSM";
                //foreach (object ob in this.UCRGrid.Columns.Select(cs => cs.Header).ToList())
                if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "25th";
                    nHeader++;
                }
                if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "30th";
                    nHeader++;
                }
                if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "35th";
                    nHeader++;
                }
                if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "340th";
                    nHeader++;
                }
                if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "45th";
                    nHeader++;
                }
                if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "50th";
                    nHeader++;
                }
                if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "55th";
                    nHeader++;
                }
                if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "60th";
                    nHeader++;
                }
                if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "65th";
                    nHeader++;
                }
                if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "70th";
                    nHeader++;
                }
                if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "75th";
                    nHeader++;
                }
                if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "80th";
                    nHeader++;
                }
                if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "85th";
                    nHeader++;
                }
                if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "90th";
                    nHeader++;
                }
                if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                {
                    m_objSheet.Cells[1, nHeader] = "95th";
                    nHeader++;
                }


                int nRow = 2;
                int nColumn;

                foreach (DataRowView field in App.myMainWin.ANEStdDataGrid.ItemsSource)
                {
                    //DataRow DR = UCRGrid.Row;
                    DataRowView DR = field;
                    m_objSheet.Cells[nRow, 1] = DR[0]; //Zipcode
                    m_objSheet.Cells[nRow, 2] = "'" + DR[1]; //CPT
                    m_objSheet.Cells[nRow, 3] = DR[2]; //Type
                    m_objSheet.Cells[nRow, 4] = "'" + DR[3]; //ANE
                    m_objSheet.Cells[nRow, 5] = DR[4]; //Status
                    m_objSheet.Cells[nRow, 6] = DR[5]; //Description
                    m_objSheet.Cells[nRow, 7] = DR[6]; //Minutes
                    m_objSheet.Cells[nRow, 8] = DR[7]; //PSM

                    nColumn = 9;
                    if (App.myMainWin.bV25th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[8];
                        nColumn++;
                    }
                    if (App.myMainWin.bV30th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[9];
                        nColumn++;
                    }
                    if (App.myMainWin.bV35th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[10];
                        nColumn++;
                    }
                    if (App.myMainWin.bV40th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[11];
                        nColumn++;
                    }
                    if (App.myMainWin.bV45th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[12];
                        nColumn++;
                    }
                    if (App.myMainWin.bV50th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[13];
                        nColumn++;
                    }
                    if (App.myMainWin.bV55th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[14];
                        nColumn++;
                    }
                    if (App.myMainWin.bV60th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[15];
                        nColumn++;
                    }
                    if (App.myMainWin.bV65th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[16];
                        nColumn++;
                    }
                    if (App.myMainWin.bV70th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[17];
                        nColumn++;
                    }
                    if (App.myMainWin.bV75th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[18];
                        nColumn++;
                    }
                    if (App.myMainWin.bV80th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[19];
                        nColumn++;
                    }
                    if (App.myMainWin.bV85th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[20];
                        nColumn++;
                    }
                    if (App.myMainWin.bV90th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[21];
                        nColumn++;
                    }
                    if (App.myMainWin.bV95th == System.Windows.Visibility.Visible)
                    {
                        m_objSheet.Cells[nRow, nColumn] = DR[22];
                        nColumn++;
                    }
                    nRow++;
                }
                m_objSheet.Columns["A:U"].AutoFit();
                m_objExcel.Visible = true;
                // Save the Workbook and quit Excel.
                //m_objBook.SaveAs("D:\\Temp\\Book2.xlsx");
                //m_objBook.Close(true, m_objOpt, false);
                //m_objExcel.Quit();
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                //m_objExcel.Quit();
                return;
            }
        }

        public static void Print_ANEBASE()
        {
            try
            {
                StringBuilder str = new StringBuilder();
                str.Append(@"&""Calibri,Bold,&14FeeViewer Pro Print");

                //Global settings
                Microsoft.Office.Interop.Excel.Application m_objExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook m_objBook = m_objExcel.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet m_objSheet = m_objBook.Worksheets.get_Item(1);

                // set header and footer settings
                m_objSheet.PageSetup.CenterHeader = str.ToString();
                m_objSheet.PageSetup.LeftFooter = App.myMainWin.OpenDbName.Text;
                m_objSheet.PageSetup.CenterFooter = "&D";
                m_objSheet.PageSetup.RightFooter = "&P/&N";

                // set body formats
                m_objSheet.PageSetup.PrintTitleRows = "$1:$1";  // repeats header row across pages
                m_objSheet.PageSetup.LeftMargin = 20;
                m_objSheet.PageSetup.RightMargin = 20;
                m_objSheet.PageSetup.TopMargin = 50;
                m_objSheet.PageSetup.BottomMargin = 50;
                m_objSheet.PageSetup.Zoom = false;
                m_objSheet.PageSetup.FitToPagesWide = 1;
                m_objSheet.PageSetup.FitToPagesTall = false;
                m_objSheet.PageSetup.PrintGridlines = true;

                m_objSheet.Cells[2, 4].EntireColumn.NumberFormat = "@";
                m_objSheet.Columns[6].ColumnWidth = 40;
                m_objSheet.Columns[9].ColumnWidth = 20;
                m_objSheet.Columns[10].ColumnWidth = 20;
                m_objSheet.Columns[11].ColumnWidth = 20;
                m_objSheet.Columns[12].ColumnWidth = 20;
                for (int i = 9; i < 13; i++)
                    m_objSheet.Cells[2, i].EntireColumn.NumberFormat = "#,##0.00";

                m_objSheet.Cells[1, 1] = "Zipcode";
                m_objSheet.Cells[1, 2] = "CPT";
                m_objSheet.Cells[1, 3] = "Type";
                m_objSheet.Cells[1, 4] = "ANE";
                m_objSheet.Cells[1, 5] = "Status";
                m_objSheet.Cells[1, 6] = "Description";
                m_objSheet.Cells[1, 7] = "Minutes";
                m_objSheet.Cells[1, 8] = "PSM";
                m_objSheet.Cells[1, 9] = "High Base High Minute";
                m_objSheet.Cells[1, 10] = "High Base Low Minute";
                m_objSheet.Cells[1, 11] = "Low Base High Minute";
                m_objSheet.Cells[1, 12] = "Low Base Low Minute";

                int nRow = 2;

                foreach (DataRowView field in App.myMainWin.ANEBaseGrid1.ItemsSource)
                {
                    //DataRow DR = UCRGrid.Row;
                    DataRowView DR = field;
                    m_objSheet.Cells[nRow, 1] = DR[0]; //Zipcode
                    m_objSheet.Cells[nRow, 2] = "'" + DR[1]; //CPT
                    m_objSheet.Cells[nRow, 3] = DR[2]; //Type
                    m_objSheet.Cells[nRow, 4] = "'" + DR[3]; //ANE
                    m_objSheet.Cells[nRow, 5] = DR[4]; //Status
                    m_objSheet.Cells[nRow, 6] = DR[5]; //Description
                    m_objSheet.Cells[nRow, 7] = DR[6]; //Minutes
                    m_objSheet.Cells[nRow, 8] = DR[7]; //PSM
                    m_objSheet.Cells[nRow, 9] = DR[8]; //High Base High Minute
                    m_objSheet.Cells[nRow, 10] = DR[9]; //High Base Low Minute
                    m_objSheet.Cells[nRow, 11] = DR[10]; //Low Base High Minute
                    m_objSheet.Cells[nRow, 12] = DR[11]; //Low Base Low Minute
                    nRow++;
                }
                m_objSheet.Columns["A:U"].AutoFit();
                m_objExcel.Visible = true;
                // Save the Workbook and quit Excel.
                //m_objBook.SaveAs("D:\\Temp\\Book2.xlsx");
                //m_objBook.Close(true, m_objOpt, false);
                //m_objExcel.Quit();
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message.ToString(), "Load Error");
                //m_objExcel.Quit();
                return;
            }
        }

    }
}
