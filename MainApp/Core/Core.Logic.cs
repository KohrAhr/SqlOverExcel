using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelWorkbookSplitter.Functions;

namespace ExcelWorkbookSplitter.Core
{
    public class LogicCore
    {
        public int ProceedFile(String fileNameIn, String fileNameOut, String worksheetName, int headerEnd, int requiredNumberOfOutputFiles)
        {
//            using (ExcelFile excelFile = new ExcelFunctions().OpenExcelFile(fileNameIn))
            {
                // Get page
                //ExcelWorksheet excelWorksheet = GetExcelWorksheet(excelFile, worksheetName);

                //// Cache header
                //IList<string> header = GetHeaderFromWorksheet(excelWorksheet, headerEnd);

                //// Calculate .......
                //int countOfRows = GetCountOfRows(excelWorksheet);
                //int recordPerFile = countOfRows;
                //if (countOfRows >= 10000)
                //{
                //    recordPerFile /= requiredNumberOfOutputFiles;
                //}


                ////if (
          
    //
            }

            return 0;
        }
    }
}
