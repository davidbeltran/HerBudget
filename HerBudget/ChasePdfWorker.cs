using System.Collections;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class ChasePdfWorker : PdfWorker
    {
        public ChasePdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = "(?:\\n((?:0[1-9]|1[1,2])/(?:0[1-9]|[12][0-9]|3[01]))\\s*(.+)" +
                " ((?:-\\d+\\.\\d{2})|(?:\\d+\\.\\d{2})))";
        }

        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();

            foreach (Match match in matches)
            {
                DateTime date = DateTime.Parse($"{match.Groups[1].Value}/{GetYear()}");
                string detail = match.Groups[2].Value.ToUpper();
                double amount = double.Parse(match.Groups[3].Value);
                if ((amount < 0) && (!Regex.IsMatch(detail, "OFFER:")))
                {
                    continue;
                }
                Expense exp = new Expense(date, detail, amount);
                if ((amount < 0) && (Regex.IsMatch(detail, "OFFER:")))
                {
                    exp.Amount = Math.Abs(amount);
                    exp.Category = CategoryType.INCOME;
                    exp.SubCategory = null;
                }
                ExpenseList.Add(exp);
            }
            IComparer comparer = (IComparer)new DateComparer();
            ExpenseList.Sort(comparer);
            foreach (Expense exp in ExpenseList)
            {
                Console.WriteLine(exp.Month);
            }
            return ExpenseList;
        }
    }
}
