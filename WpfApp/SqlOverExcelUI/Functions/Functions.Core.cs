using System;
using System.Collections.Generic;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelObject = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace SqlOverExcelUI.Functions
{
    public class CoreFunctions
    {
        public void SaveResultToExcelFile(string toFile, DataTable data)
        {
            // Save result to new file
            using (ExcelCore excelOut = new ExcelCore())
            {
                excelOut.NewFile(toFile);
                if (excelOut.IsInitialized())
                {
                    excelOut.NewSheet("RESULT", WorksheerOrder.woFirst);

                    // Delete default worksheet
                    excelOut.DeleteSheet("Sheet1");

                    excelOut.PopulateData("RESULT", data);

                    excelOut.SaveFile();
                }
            }
        }
    }
}
