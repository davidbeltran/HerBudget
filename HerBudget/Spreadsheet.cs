using System.Collections;
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

    }
}
