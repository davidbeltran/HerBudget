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
                double amount = double.Parse(match.Groups[3].Value);
                if (amount < 0)
                {
                    continue;
                }
                var temp = new ArrayList();
                temp.Add(DateTime.Parse($"{match.Groups[1].Value}/{GetYear()}"));
                temp.Add(match.Groups[2].Value.ToUpper());
                temp.Add(amount);
                ExpenseList.Add(temp);
            }
            foreach (ArrayList exp in  ExpenseList)
            {
                Console.WriteLine($"Date: {exp[0]} || Detail1:{exp[1]} || Amount: {exp[2]}\n");
            }
            return ExpenseList;
        }
    }
}
