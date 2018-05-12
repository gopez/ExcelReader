using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename    = "data.xlsx";
            string projectPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName; 
            string dataPath    = Path.Combine(projectPath, filename);
            
            Excel.Application xlApp      = new Excel.Application();
            Excel.Workbook xlWorkbook    = xlApp.Workbooks.Open(Path.Combine(projectPath, dataPath));
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange          = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // data starts from row = 2
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 1)
                        Console.Write("\r\n");

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    }
                }
            }

            // cleaning-up
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine();
            Console.ReadKey();
        }
    }
}
