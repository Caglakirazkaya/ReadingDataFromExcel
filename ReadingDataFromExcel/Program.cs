using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadingDataFromExcel
{
    class Program
    {
        static void Main(string[] args)
        { 
            Console.Write("Start \n");

            Application excelApp = new Application();
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\ck\Desktop\ExcelReading\ExcelRead.xlsx");

            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;          

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {                   
                    Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
                }
                
                Console.Write("\n");
            }
            Console.Write("End \n");
            //after reading, relaase the excel project
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadLine();

        }
    }
}
