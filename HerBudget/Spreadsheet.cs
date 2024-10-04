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
            Expense firstExp = (Expense)this.Expenses[0]!;
            Expense lastExp = (Expense)this.Expenses[^1]!;
            string fullPath = MakeDirectory() + @"\Finances" + firstExp.Year + ".xlsx";
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;

            //TODO
            //-still need to adjust AddBills() method to separte adding amounts to their appropriate months
            //-ai came up with this solution below: 
            //      Excel.Range cell = worksheet.Cells[1, 1]; // Row 1, Column 1 (A1)
            //      string cellValue = cell.Value.ToString();
            //-look at this ai solution to perhaps take care of the task manager issue
            //      Excel.Application excelApp = new Excel.Application();
            //      Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\path\to\your\excel\file.xlsx");
            //      workbook.Close(false);
            //      excelApp.Quit();
            if (!File.Exists(fullPath))
            {
                Excel.Workbook workbook = excel.Workbooks.Add();
                Excel._Worksheet worksheet = (Excel.Worksheet)excel.ActiveSheet;
                worksheet.Name = firstExp.Month;
                AddBills(worksheet);

                if (firstExp.Month != lastExp.Month)
                {
                    worksheet = (Excel.Worksheet)workbook.Sheets.Add(workbook.Sheets[workbook.Sheets.Count],
                        System.Type.Missing, System.Type.Missing, System.Type.Missing);
                    worksheet.Name = lastExp.Month;
                    AddBills(worksheet);
                }
                workbook.SaveAs(fullPath);
                workbook.Close(false);
                excel.Quit();
            }
            else
            {
                Excel.Workbook workbook = excel.Workbooks.Open(fullPath);
                Excel.Sheets worksheets = workbook.Sheets;

                AddBothMonths(worksheets, firstExp, lastExp);
                workbook.Save();
                workbook.Close(false);
                excel.Quit();
            }
        }

        private void AddBothMonths(Excel.Sheets sheets, Expense firstExp, Expense lastExp)
        {
            AddMonth(sheets, firstExp);
            if (firstExp.Month != lastExp.Month)
            {
                AddMonth(sheets, lastExp);
            }
        }

        private void AddMonth(Excel.Sheets sheets, Expense exp)
        {
            Excel._Worksheet worksheet;
            if (FindSheet(sheets, exp))
            {
                worksheet = (Excel.Worksheet)sheets[exp.Month];
            }
            else
            {
                worksheet = (Excel.Worksheet)sheets.Add(sheets[sheets.Count],
                    System.Type.Missing, System.Type.Missing, System.Type.Missing);
                worksheet.Name = exp.Month;
            }
            AddBills(worksheet);
        }
        private bool FindSheet(Excel.Sheets sheets, Expense exp)
        {
            foreach (Excel.Worksheet sheet in sheets)
            {
                if (sheet.Name == exp.Month)
                {
                    return true;
                }
            }
            return false;
        }

        private void AddBills(Excel._Worksheet sheet)
        {
            double Internet = 0, Car_Insurance = 0, Housing = 0, Energy = 0, Gas = 0, Income = 0,
                Phones = 0, Entertainment = 0, Dental = 0, Healthcare = 0, Savings = 0, Car_Gas = 0,
                Groceries = 0, Restaurant = 0, Necessary = 0, Unnecessary = 0;
            DisplayInstructions();
            foreach (Expense exp in  Expenses)
            {
                if (!exp.Category.Equals(CategoryType.EXPENSE))
                {
                    switch (exp.SubCategory)
                    {
                        case SubCategoryType.INTERNET: Internet += exp.Amount; break;
                        case SubCategoryType.CAR_INSURANCE: Car_Insurance += exp.Amount; break;
                        case SubCategoryType.RENT: Housing += exp.Amount; break;
                        case SubCategoryType.ELECTRIC: Energy += exp.Amount; break;
                        case SubCategoryType.GAS_HOME: Gas += exp.Amount; break;
                        case SubCategoryType.PHONES: Phones += exp.Amount; break;
                        case SubCategoryType.TV: Entertainment += exp.Amount; break;
                        case SubCategoryType.DENTAL: Dental += exp.Amount; break;
                        case SubCategoryType.HEALTHCARE: Healthcare += exp.Amount; break;
                        case SubCategoryType.SAVINGS: Savings += exp.Amount; break;
                        default: Income += exp.Amount; break;
                    }
                }
                else
                {
                    Expense newExp = AskUser(exp);
                    switch(exp.SubCategory)
                    {
                        case SubCategoryType.GAS_CAR: Car_Gas += exp.Amount; break;
                        case SubCategoryType.GROCERIES: Groceries += exp.Amount; break;
                        case SubCategoryType.RESTAURANT: Restaurant += exp.Amount; break;
                        case SubCategoryType.MISC_NECESSARY: Necessary += exp.Amount; break;
                        case SubCategoryType.MISC_UNNECESSARY: Unnecessary += exp.Amount; break;
                    }
                }
            }

            MakeHeaders(sheet);

            //BILLS
            sheet.Cells[2, 3] = Internet;
            sheet.Cells[3, 3] = Car_Insurance;
            sheet.Cells[4, 3] = Housing;
            sheet.Cells[5, 3] = Energy;
            sheet.Cells[6, 3] = Gas;
            sheet.Cells[7, 3] = Phones;
            sheet.Cells[8, 3] = Entertainment;
            sheet.Cells[9, 3] = Dental;
            sheet.Cells[10, 3] = Healthcare;
            sheet.Cells[11, 3] = Savings;
            sheet.Cells[12, 3] = "=SUM(C2:C11)";

            //EXPENSES
            sheet.Cells[15, 3] = Car_Gas;
            sheet.Cells[16, 3] = Groceries;
            sheet.Cells[17, 3] = Restaurant;
            sheet.Cells[18, 3] = Necessary;
            sheet.Cells[19, 3] = Unnecessary;
            sheet.Cells[20, 3] = "=SUM(C15:C19)";

            //CALCULATIONS
            sheet.Cells[22, 3] = "=SUM(C20,C12)";
            sheet.Cells[24, 3] = Income;
            sheet.Cells[26, 3] = "=C24-C22";
        }

        private Expense AskUser(Expense exp)
        {
            string? ResponseCheck;
            while (true)
            {
                string selection = DisplayTransaction(exp);
                Console.WriteLine($"You selected {selection}.\nIs this correct?");
                ResponseCheck = Console.ReadLine();
                if (ResponseCheck!.ToLower() == "n" || ResponseCheck!.ToLower() == "no")
                {
                    continue;
                }
                else
                {
                    switch(selection)
                    {
                        case "Gas": exp.SubCategory = SubCategoryType.GAS_CAR; break;
                        case "Groceries": exp.SubCategory = SubCategoryType.GROCERIES; break;
                        case "Restaurant": exp.SubCategory = SubCategoryType.RESTAURANT; break;
                        case "Misc(Necessary)": exp.SubCategory = SubCategoryType.MISC_NECESSARY; break;
                        case "Misc(Unnecessary)": exp.SubCategory= SubCategoryType.MISC_UNNECESSARY; break;
                    }
                    break;
                }
            }
            return exp;
        }

        private void DisplayInstructions()
        {
            Console.WriteLine("All bill transactions are sorted. Please sort expense transactions.\n" +
                                "Select number corresponding to category.\n" +
                                "1. Gas\n2. Groceries\n3. Restaurant\n4. Misc(Necessary)\n5. Misc(Unnecessary)");
        }

        private string DisplayTransaction(Expense exp)
        {
            Console.WriteLine("\n\n\n" +
                "Transaction\n===========================================================================================================\n" +
                $"Date: {exp.Date.ToShortDateString()}   |   Detail: {exp.Detail}   |   Amount: ${exp.Amount}\n" +
                "===========================================================================================================\n" +
                "1. Gas, 2. Groceries, 3. Restaurant, 4. Misc(Necessary), or 5. Misc(Unnecessary)?");
            string? catSelection;
            int catSelectionNum;
            while (true)
            {
                catSelection = Console.ReadLine();
                if (int.TryParse(catSelection, out catSelectionNum) && catSelectionNum >= 1 && catSelectionNum <= 5)
                {
                    break;
                }
                else
                {
                    Console.WriteLine("Selection must be a number between 1 and 5. Please try again:");
                    continue;
                }
            }
            switch(catSelectionNum)
            {
                case 1: catSelection = "Gas"; break;
                case 2: catSelection = "Groceries"; break;
                case 3: catSelection = "Restaurant"; break;
                case 4: catSelection = "Misc(Necessary)"; break;
                case 5: catSelection = "Misc(Unnecessary)"; break;
            }
            return catSelection;
        }
        /// <summary>
        /// Default headers created
        /// [Row, Column]
        /// </summary>
        /// <param name="sheet">Excel sheet</param>
        private static void MakeHeaders(Excel._Worksheet sheet)
        {
            //Bold Headers 
            sheet!.Cells[1, 1] = "TYPE";
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
