using Microsoft.Office.Interop.Excel;
using System.Collections;
using UglyToad.PdfPig.Graphics.Operations.PathPainting;
using Excel = Microsoft.Office.Interop.Excel;

namespace HerBudget
{
    public class Spreadsheet
    {
        public ArrayList Expenses { get; set; }

        public Spreadsheet(ArrayList expenses)
        {
            this.Expenses = expenses;
        }

        private static string MakeDirectory()
        {
            string sheetPath = Directory.GetParent(Environment.CurrentDirectory)!.Parent!.FullName
                + @"\HerBudget\sheets";
            if (!Directory.Exists(sheetPath))
            {
                Directory.CreateDirectory(sheetPath);
            }
            return sheetPath;
        }

        public void AddToExcel()
        {
            string fullPath = MakeDirectory() + @"\Finances.xlsx";
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;

            Excel.Workbook workbook = excel.Workbooks.Add();
            Excel._Worksheet worksheet = (Excel.Worksheet)excel.ActiveSheet;
            MakeHeaders(worksheet);
            workbook.SaveAs(fullPath);
            workbook.Close();

            Console.WriteLine(fullPath);
        }

        private static void MakeHeaders(Excel._Worksheet sheet)
        {
            sheet.Cells[1, 1] = "TYPE";
        }
    }
}
