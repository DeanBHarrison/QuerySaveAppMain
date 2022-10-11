using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
// USE EXCEL OBJECT LIBRARY
using System.Reflection;
// USE SYSTEM.DATA NOT MICROSOFT.DATA
using System.Data.SqlClient;
using System.Data;

namespace QuerySaveApp
{
    public class ExcelWrite
    {
        public static void xlsWorkBook(string StoredProcedureName, string SaveLocation)
        {
            Excel.Application oXL;
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;
            Excel.Range oRange;

            // Start Excel and get Application object.
            oXL = new Excel.Application();
            // Set some properties
            oXL.Visible = false;
            oXL.DisplayAlerts = false;

            // Get a new workbook.
            oWB = oXL.Workbooks.Add(Missing.Value);

           // oWB = oXL.Workbooks.Open(@"\\SRVLICHFDFP1\Asos$\MI\MI Reporting\136 Order Consolidation\Order Consolidation Lichfield.xlsx");
            // Get the active sheet
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;
           // oSheet.Name = "Customers";
            // Process the DataTable
            //this is calling 
            DataTable dt = SQLAcess.SQLtoDataTable(StoredProcedureName);//commented
                                                               // DataTable dt = Table;//added
                                                               //Session["dt"] = dt;//added
            int rowCount = 1;
            foreach (DataRow dr in dt.Rows)
            {

                for (int i = 1; i < dt.Columns.Count + 1; i++)
                {

                    oSheet.Cells[rowCount, i] = dr[i - 1]
                        .ToString();
                }
                rowCount += 1;
            }

            // Resize the columns
           // oRange = oSheet.get_Range(oSheet.Cells[1, 1],
           // oSheet.Cells[rowCount, dt.Columns.Count]);
           // oRange.EntireColumn.AutoFit();
            // Save the sheet and close
           //oSheet = null;
           //oRange = null;
            oWB.SaveAs(SaveLocation + @".xlsx");
            oWB.Close();
           // oWB = null;
            oXL.Quit();
            // Clean up
            // NOTE: When in release mode, this does the trick
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
