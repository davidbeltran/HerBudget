using System.Collections;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class AllyPdfWorker : PdfWorker
    {
        public AllyPdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = @"((?:\n0[1-9]|1[1,2])/(?:0[1-9]|[1-2][0-9]|3[0-1])/(?:\d{4})).*" +
                "(\n.*\n)";
        }
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();
            foreach (Match match in matches)
            {
                Console.WriteLine(match.Groups[1] + " " + match.Groups[2]);
            }
            return ExpenseList;
        }
    }
}
