/// WARNING: ANY USE BY YOU OF THE SAMPLE CODE PROVIDED IN THIS FILE IS AT YOUR OWN RISK.
/// Microsoft provides this code "as is" without warranty of any kind, either express or implied, 
/// including but not limited to the implied warranties of merchantability and/or fitness 
/// for a particular purpose.

using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAutoFitSample
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateAutoFitExcelFile();
            Console.ReadLine();
        }

        public static void CreateAutoFitExcelFile()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[1, 3] = "Address";
            xlWorkSheet.Cells[2, 1] = "abcdefghigklmn";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[2, 3] = "Three";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "abcdefghigklmn";
            xlWorkSheet.Cells[3, 3] = "cccccc";

            //     Changes the width of the columns in the range or the height of the rows in the
            //     range to achieve the best fit.
            xlWorkSheet.Columns.AutoFit();

            //xlWorkSheet.Rows.AutoFit();

            string excelFilePath = @"C:\temp\csharp-Excel.xlsx";

            xlApp.ActiveWorkbook.SaveCopyAs(excelFilePath);
            xlApp.ActiveWorkbook.Saved = true;

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Excel file created , you can find the file " + excelFilePath);

        }
    }
}
