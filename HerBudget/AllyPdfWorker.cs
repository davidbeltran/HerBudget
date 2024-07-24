using System.Collections;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class AllyPdfWorker : PdfWorker
    {
        public AllyPdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = @"\n((?:0[1-9]|1[1,2])/(?:0[1-9]|[1-2][0-9]|3[0-1])/(?:\d{4})) " +
                @"(?:Check Card Purchase\n|ACH Withdrawal\n|Direct Deposit\n|)";
        }
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();
            foreach (Match match in matches)
            {
                Console.WriteLine($"Date: {match.Groups[1]} || Detail:{match.Groups[2]} || Amount: {match.Groups[3]}");
            }
            return ExpenseList;
        }
    }
}
// Check Card Purchase
// ACH Withdrawal
// Direct Deposit
// WEB Funds Transfer
// NOW Withdrawal
// NOW Deposit