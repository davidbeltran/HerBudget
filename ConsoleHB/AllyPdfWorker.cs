/*
 * Author: David Beltran
 */

using System;
using System.Collections;
using System.Text.RegularExpressions;

namespace ConsoleHB
{
    /// <summary>
    /// Subclass of the PdfWorker abstract class for Ally Bank
    /// </summary>
    public class AllyPdfWorker : PdfWorker
    {
        /// <summary>
        /// Subclass constructor
        /// </summary>
        /// <param name="fileStorage">XML file path holding list of PDF file names</param>
        /// <param name="pdfDoc">Ally PDF file path</param>
        public AllyPdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = "(?:((?:0[1-9]|1[0-2])/(?:0[1-9]|[1-2][0-9]|3[0-1])/(?:\\d{4})) " +
            "(Check Card Purchase|ACH Withdrawal|Direct Deposit|Interest Paid|" +
            "WEB Funds Transfer|NOW Withdrawal|NOW Deposit|eCheck Deposit)\\s" +
            "(\\n.*\\s)?(?:.*\\s)*?\\$(.+\\.\\d{2}) -\\$(.+\\.\\d{2}) (?:-?\\$.+\\.\\d{2}[\\s|A]))";
        }

        /// <summary>
        /// Inherited method to create Expense object list catered specifically for Ally Bank PDF statements
        /// </summary>
        /// <returns>ArrayList of Expense objects</returns>
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList Expenses = new ArrayList();

            foreach (Match match in matches)
            {
                DateTime date = DateTime.Parse(match.Groups[1].Value.Trim());
                string detail1 = match.Groups[2].ToString().Trim().ToUpper();
                string detail2 = match.Groups[3].ToString().Trim().ToUpper();
                double amount1 = double.Parse(match.Groups[4].Value.Trim());
                double amount2 = double.Parse((match.Groups[5].Value.Trim()));

                string detail;
                //May need to add 'or' statement to include unseen transfers from accounts outside Ally
                if ((detail2.Equals("REQUESTED TRANSFER FROM ALLY BANK")) ||
                    (detail2.Equals("CHASE CREDIT CRD EPAY~ FUTURE")) ||
                    (detail2.Equals("REQUESTED TRANSFER TO ALLY BANK SAVINGS")))
                {
                    continue; //These are not registered per client's request
                }
                else if (detail2.Equals(""))
                {
                    detail = detail1; //Takes detail if only one line exists on PDF file
                } 
                else
                {
                    detail = detail2; //Most needed detail comes from the second line on PDF file
                }

                double amount;
                if (amount1.Equals(0))
                {
                    amount = amount2; //Registers the debit amount
                }
                else
                {
                    amount = amount1; //Registers the credit amount
                }

                Expense exp = new Expense(date, detail, amount);
                if (amount.Equals(amount1)) //All credits are considered income to be added into money made
                {
                    exp.Category = CategoryType.INCOME;
                }
                Expenses.Add(exp);
            }

            IComparer comparer = new DateComparer(); //Sorts the Expenses ArrayList by Expense object month
            Expenses.Sort(comparer); //Needed to help split PDF file tables that include two months

            return Expenses;
        }
    }
}