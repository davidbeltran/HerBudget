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
        }

        /// <summary>
        /// Default headers created
        /// [Row, Column]
        /// </summary>
        /// <param name="sheet">Excel sheet</param>
        private static void MakeHeaders(Excel._Worksheet sheet)
        {
            //Bold Headers 
            sheet.Cells[1, 1] = "TYPE";
            sheet.Cells[1, 2] = "BILLS";
            sheet.Cells[1, 3] = "AMOUNT";
            sheet.Cells[12, 2] = "Total";
            sheet.Cells[14, 2] = "EXPENSES";
            sheet.Cells[20, 2] = "Total";
            sheet.Cells[22, 2] = "TOTAL SPENT";
            sheet.Cells[24, 2] = "INCOME";
            sheet.Cells[26, 2] = "CASH FLOW";
            sheet.Range["A1:C1,B12,B14,B20,B22,B24,B26"].Font.Bold = true;

            //Type labels
            sheet.Cells[2, 1] = "Internet";
            sheet.Cells[3, 1] = "Car Insurance";
            sheet.Cells[4, 1] = "Housing";
            sheet.Cells[5, 1] = "Energy";
            sheet.Cells[6, 1] = "Gas";
            sheet.Cells[7, 1] = "Phones";
            sheet.Cells[8, 1] = "Entertainment";
            sheet.Cells[9, 1] = "Dental";
            sheet.Cells[10, 1] = "Healthcare";
            sheet.Cells[11, 1] = "Savings";

            //Bill labels
            sheet.Cells[2, 2] = "Spectrum";
            sheet.Cells[3, 2] = "Progressive";
            sheet.Cells[4, 2] = "Rent";
            sheet.Cells[5, 2] = "SCE";
            sheet.Cells[6, 2] = "SoCal Gas";
            sheet.Cells[7, 2] = "AT&T";
            sheet.Cells[8, 2] = "TV";
            sheet.Cells[9, 2] = "Delta";
            sheet.Cells[10, 2] = "Kaiser";
            sheet.Cells[11, 2] = "Ally";

            //Expense labels
            sheet.Cells[15, 2] = "Car Gas";
            sheet.Cells[16, 2] = "Groceries";
            sheet.Cells[17, 2] = "Restaurants";
            sheet.Cells[18, 2] = "Misc. (Necessary)";
            sheet.Cells[19, 2] = "Misc. (Unnecessary)";
        }
    }
}
