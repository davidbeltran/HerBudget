/*
 * Author: David Beltran
 */

using System;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;

namespace ConsoleHB
{
    /// <summary>
    /// Subclass of the PdfWorker abstract class for Chase Bank
    /// </summary>
    public class ChasePdfWorker : PdfWorker
    {
        /// <summary>
        /// Subclass constructor
        /// </summary>
        /// <param name="fileStorage">XML file path holding list of PDF file names</param>
        /// <param name="pdfDoc">Chase PDF file path</param>
        public ChasePdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            //this.ReDetail = "(?:\\n((?:0[1-9]|1[0-2])/(?:0[1-9]|[12][0-9]|3[01]))\\s*(.+)" +
            //    " ((?:-\\d+\\.\\d{2})|(?:\\d+\\.\\d{2})))";
            //this.ReDetail = @"(\d{2}/\d{2})\s+(.+?)\s+(-?\d+\.\d{2})";
            //this.ReDetail = @"(\d{2}/\d{2})\s+(.+?)\s+(-?\d*\.?\d{2})";
            this.ReDetail = "(?:\\n((?:0[1-9]|1[0-2])/(?:0[1-9]|[12][0-9]|3[01]))\\s*(.+)" +
                " ((?:-\\d+\\.\\d{2})|(?:\\d*\\.\\d{2})))";
        }

        /// <summary>
        /// Inherited method to create Expense object list catered specifically for Chase Bank PDF statements
        /// </summary>
        /// <returns>ArrayList of Expense objects</returns>
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            Console.WriteLine(pdfText);
            PathCreator pc = new PathCreator("tester", "test.txt");
            string pathPdf = pc.MakeFile();
            File.WriteAllText(pathPdf, pdfText);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();

            foreach (Match match in matches)
            {
                DateTime date = DateTime.Parse($"{match.Groups[1].Value}/{GetYear()}");
                string detail = match.Groups[2].Value.ToUpper();
                double amount = double.Parse(match.Groups[3].Value);
                Console.WriteLine($"MATCH: {detail} | mAmount: {amount}");
                if ((amount < 0) && (!Regex.IsMatch(detail, "OFFER:")))
                {
                    continue; //Unregistered credits paid from Ally/checking account to credit card account
                }
                Expense exp = new Expense(date, detail, amount);
                if ((amount < 0) && (Regex.IsMatch(detail, "OFFER:"))) //Credits that come from outside client expense circulation
                {
                    exp.Amount = Math.Abs(amount);
                    exp.Category = CategoryType.INCOME;
                    exp.SubCategory = null;
                }
                ExpenseList.Add(exp);
            }
            IComparer comparer = new DateComparer(); //Sorts the Expenses ArrayList by Expense object month
            ExpenseList.Sort(comparer); //Needed to help split PDF file tables that include two months
            return ExpenseList;
        }
    }
}
