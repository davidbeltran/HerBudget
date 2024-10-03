using System.Collections;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class AllyPdfWorker : PdfWorker
    {
        public AllyPdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = "(?:((?:0[1-9]|1[1,2])/(?:0[1-9]|[1-2][0-9]|3[0-1])/(?:\\d{4})) " +
            "(Check Card Purchase|ACH Withdrawal|Direct Deposit|Interest Paid|" +
            "WEB Funds Transfer|NOW Withdrawal|NOW Deposit|eCheck Deposit)\\s" +
            "(\\n.*\\s)?(?:.*\\s)*?\\$(.+\\.\\d{2}) -\\$(.+\\.\\d{2}) (?:-?\\$.+\\.\\d{2}[\\s|A]))";
        }
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

                //May need to add or statement to include unseen transfers from accounts outside Ally
                if (detail2.Equals("REQUESTED TRANSFER FROM ALLY BANK"))
                {
                    continue;
                }
                else if (detail2.Equals(""))
                {
                    detail = detail1;
                } 
                else
                {
                    detail = detail2;
                }

                double amount;
                if (amount1.Equals(0))
                {
                    amount = amount2;
                }
                else
                {
                    amount = amount1;
                }
                Expense exp = new Expense(date, detail, amount);
                if (amount.Equals(amount1))
                {
                    exp.Category = CategoryType.INCOME;
                }
                Expenses.Add(exp);
            }
            IComparer comparer = new DateComparer();
            Expenses.Sort(comparer);
            return Expenses;
        }
    }
}