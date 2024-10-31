/*
 * Author: David Beltran
 */

using System;
using System.Collections;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleHB
{
    /// <summary>
    /// Class that holds information for each transaction
    /// </summary>
    public class Spreadsheet
    {
        public ArrayList Expenses { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="expenses">ArrayList to hold Expense objects</param>
        public Spreadsheet(ArrayList expenses)
        {
            this.Expenses = expenses;
        }

        /// <summary>
        /// Loads Expense object data to Excel workbook
        /// </summary>
        public void AddToExcel()
        {
            Expense firstExp = (Expense)this.Expenses[0]!; //First Expense object of list
            Expense lastExp = (Expense)this.Expenses[^1]!; //Last Expense object of list
            PathCreator pc = new PathCreator("sheets", $"Finances{firstExp.Year}.xlsx");
            string fullPath = pc.MakeFile();
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;

            if (!File.Exists(fullPath)) //Completely new workbook
            {
                Excel.Workbook workbook = excel.Workbooks.Add();
                Excel._Worksheet worksheet = (Excel.Worksheet)excel.ActiveSheet;
                worksheet.Name = firstExp.Month;
                AddBills(worksheet);

                if (firstExp.Month != lastExp.Month) //new month/sheet on existing workbook
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
            else //adding data to existing worksheet months
            {
                Excel.Workbook workbook = excel.Workbooks.Open(fullPath);
                Excel.Sheets worksheets = workbook.Sheets;

                AddBothMonths(worksheets, firstExp, lastExp);
                workbook.Save();
                workbook.Close(false);
                excel.Quit();
            }
        }

        /// <summary>
        /// Allows two months to be added. Used in AddToExcel() method
        /// </summary>
        /// <param name="sheets">Existing workbook months</param>
        /// <param name="firstExp">First Expense object of list</param>
        /// <param name="lastExp">Last Expense object of list</param>
        private void AddBothMonths(Excel.Sheets sheets, Expense firstExp, Expense lastExp)
        {
            AddMonth(sheets, firstExp);
            if (firstExp.Month != lastExp.Month)
            {
                AddMonth(sheets, lastExp);
            }
        }

        /// <summary>
        /// Adds one month at a time by checking if month exists in workbook
        /// </summary>
        /// <param name="sheets">Existing workbook months</param>
        /// <param name="exp">Current Expense object</param>
        private void AddMonth(Excel.Sheets sheets, Expense exp)
        {
            Excel._Worksheet worksheet;
            if (FindSheet(sheets, exp)) //curent sheet selected as existing month
            {
                worksheet = (Excel.Worksheet)sheets[exp.Month];
            }
            else //creates a new worksheet to add to workbook
            {
                worksheet = (Excel.Worksheet)sheets.Add(sheets[sheets.Count],
                    System.Type.Missing, System.Type.Missing, System.Type.Missing);
                worksheet.Name = exp.Month;
            }
            AddBills(worksheet);
        }

        /// <summary>
        /// Checks if sheet exists in workbook worksheets
        /// </summary>
        /// <param name="sheets">Existing workbook months</param>
        /// <param name="exp">Current Expense object</param>
        /// <returns>True if month already exists in workbook.</returns>
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

        /// <summary>
        /// Registers Expense object data to excel worksheet
        /// </summary>
        /// <param name="sheet">Current worksheet/month</param>
        private void AddBills(Excel._Worksheet sheet)
        {
            double Internet, Car_Insurance, Housing, Energy, Gas, Income,
                Phones, Entertainment, Dental, Healthcare, Savings, Car_Gas,
                Groceries, Restaurant, Necessary, Unnecessary;

            FindCellValues(out Internet, out Car_Insurance, out Housing, out Energy, out Gas,
                out Income, out Phones, out Entertainment, out Dental, out Healthcare,
                out Savings, out Car_Gas, out Groceries, out Restaurant,
                out Necessary, out Unnecessary, sheet);
                        
            DisplayInstructions();

            foreach (Expense exp in  Expenses)
            {
                if (exp.Month == sheet.Name) //Organizes Expense month into proper worksheet month
                {
                    if (!exp.Category.Equals(CategoryType.EXPENSE)) //Automatically excel registered Expense objects
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
                            //case SubCategoryType.SAVINGS: Savings += exp.Amount; break;
                            default: Income += exp.Amount; break;
                        }
                    }
                    else //Requires user clarification on subcategory of Expense object
                    {
                        AskUser(exp);
                        switch (exp.SubCategory)
                        {
                            case SubCategoryType.GAS_CAR: Car_Gas += exp.Amount; break;
                            case SubCategoryType.GROCERIES: Groceries += exp.Amount; break;
                            case SubCategoryType.RESTAURANT: Restaurant += exp.Amount; break;
                            case SubCategoryType.MISC_NECESSARY: Necessary += exp.Amount; break;
                            case SubCategoryType.MISC_UNNECESSARY: Unnecessary += exp.Amount; break;
                        }
                    }
                }
                else
                {
                    continue;
                }
            }

            MakeHeaders(sheet); //Labels and headers added to worksheet first

            //Expense amount total registered to excel worksheet
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
            //sheet.Cells[11, 3] = Savings;
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

        /// <summary>
        /// Scans current worksheet for possible existing amounts to be added to Expense objects with same months from different banks.
        /// </summary>
        /// <param name="Internet">Expense subcategory</param>
        /// <param name="Car_Insurance">Expense subcategory</param>
        /// <param name="Housing">Expense subcategory</param>
        /// <param name="Energy">Expense subcategory</param>
        /// <param name="Gas">Expense subcategory</param>
        /// <param name="Income">Expense subcategory</param>
        /// <param name="Phones">Expense subcategory</param>
        /// <param name="Entertainment">Expense subcategory</param>
        /// <param name="Dental">Expense subcategory</param>
        /// <param name="Healthcare">Expense subcategory</param>
        /// <param name="Savings">Expense subcategory</param>
        /// <param name="Car_Gas">Expense subcategory</param>
        /// <param name="Groceries">Expense subcategory</param>
        /// <param name="Restaurant">Expense subcategory</param>
        /// <param name="Necessary">Expense subcategory</param>
        /// <param name="Unnecessary">Expense subcategory</param>
        /// <param name="sheet">Current excel worksheet</param>
        private void FindCellValues(out double Internet, out double Car_Insurance, out double Housing,
            out double Energy, out double Gas, out double Income, out double Phones, out double Entertainment,
            out double Dental, out double Healthcare, out double Savings, out double Car_Gas, out double Groceries,
            out double Restaurant, out double Necessary, out double Unnecessary, Excel._Worksheet sheet)
        {
            Internet = CellValue(2, 3, sheet);
            Car_Insurance = CellValue(3, 3, sheet);
            Housing = CellValue(4, 3, sheet);
            Energy = CellValue(5, 3, sheet);
            Gas = CellValue(6, 3, sheet);
            Income = CellValue(24, 3, sheet);
            Phones = CellValue(7, 3, sheet);
            Entertainment = CellValue(8, 3, sheet);
            Dental = CellValue(9, 3, sheet);
            Healthcare = CellValue(10, 3, sheet);
            //Savings = CellValue(11, 3, sheet);
            Car_Gas = CellValue(15, 3, sheet);
            Groceries = CellValue(16, 3, sheet);
            Restaurant = CellValue(17, 3, sheet);
            Necessary = CellValue(18, 3, sheet);
            Unnecessary = CellValue(19, 3, sheet);
        }

        /// <summary>
        /// Scans given cell for possible existing value
        /// </summary>
        /// <param name="row">Excel row</param>
        /// <param name="column">Excel column</param>
        /// <param name="sheet">Current worksheet</param>
        /// <returns></returns>
        private double CellValue(int row, int column, Excel._Worksheet sheet)
        {
            double Value;
            Excel.Range CellVal = (Excel.Range)sheet.Cells[row, column];

            if (CellVal.Value != null)
            {
                Value = Convert.ToDouble(CellVal.Value.ToString());
            }
            else
            {
                Value = 0.0;
            }
            return Value;
        }

        /// <summary>
        /// UI asking user to provide Subcategory of each Expense object
        /// </summary>
        /// <param name="exp">Current Expense object</param>
        private void AskUser(Expense exp)
        {
            string ResponseCheck;
            while (true)
            {
                string selection = DisplayTransaction(exp);
                Console.WriteLine($"You selected {selection}.\nIs this correct? (Y/N)");
                ResponseCheck = Console.ReadLine()!.Trim().ToLower();
                if (ResponseCheck == "y" || ResponseCheck == "yes")
                {
                    switch (selection)
                    {
                        case "Gas": exp.SubCategory = SubCategoryType.GAS_CAR; break;
                        case "Groceries": exp.SubCategory = SubCategoryType.GROCERIES; break;
                        case "Restaurant": exp.SubCategory = SubCategoryType.RESTAURANT; break;
                        case "Misc(Necessary)": exp.SubCategory = SubCategoryType.MISC_NECESSARY; break;
                        case "Misc(Unnecessary)": exp.SubCategory = SubCategoryType.MISC_UNNECESSARY; break;
                    }
                    break;
                }
                else if (ResponseCheck == "n" || ResponseCheck == "no")
                {
                    Console.WriteLine("\nNo problem :)\nGo ahead and try again.");
                    continue;
                }
                else
                {
                    Console.WriteLine("\n\nYou must select (y)es or (n)o only.");
                    continue;
                }
            }
        }

        /// <summary>
        /// Displays AskUser() UI instructions for user.
        /// </summary>
        private void DisplayInstructions()
        {
            Console.WriteLine("\n\n\n\n\n\n\n\n\n\n\n*******************************************************************\n" + 
                "All bill transactions are sorted. Please sort expense transactions.\n" +
                "Select number corresponding to category.\n" +
                "*******************************************************************\n" +
                "1. Gas\n2. Groceries\n3. Restaurant\n4. Misc(Necessary)\n5. Misc(Unnecessary)");
        }

        /// <summary>
        /// Display each Expense object's details of transactions needing user subcategory selection
        /// </summary>
        /// <param name="exp">Current Expense object</param>
        /// <returns>string value of user subcategory selection</returns>
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
        /// Default headers created in current Excel worksheet
        /// [Row, Column]
        /// </summary>
        /// <param name="sheet">Current Excel sheet</param>
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
            //sheet.Cells[11, 1] = "Savings";

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
