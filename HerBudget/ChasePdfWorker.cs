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
            Console.WriteLine(pdfText);
            //MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            //ArrayList ExpenseList = new ArrayList();
            //string skip = "Payment Thank You-Mobile";
            //foreach (Match match in matches)
            //{
            //    if (match.Groups[2].Value == skip)
            //    {
            //        continue;
            //    }
            //    var temp = new ArrayList();
            //    temp.Add(DateTime.Parse($"{match.Groups[1].Value}/{GetYear()}"));
            //    temp.Add(match.Groups[2].Value);
            //    temp.Add(double.Parse(match.Groups[3].Value));
            //    ExpenseList.Add(temp);
            //}
            //return ExpenseList;
            return new ArrayList();
        }
    }
}
