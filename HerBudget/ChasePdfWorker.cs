using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig;
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

        /// <summary>
        /// Retrieves third page of pdf and converts to .txt file
        /// </summary>
        /// <param name="pdfDoc"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public override string PreparePdf(string pdfDoc)
        {
            string PageText = "";
            try
            {
                using (PdfDocument doc = PdfDocument.Open(pdfDoc))
                {
                    PageText = ContentOrderTextExtractor.GetText(doc.GetPage(3));
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException("**PDF file not included in folder. Copy from Python project**", ex.Message);
            }
            return PageText;
        }

        /// <summary>
        /// Finds date, detail, and amount of each expense.
        /// Adds the data to an arraylist.
        /// Parses string date into DateTime type and string amount into double type.
        /// </summary>
        /// <returns>ArrayList of expense details</returns>
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();
            string skip = "Payment Thank You-Mobile";
            foreach (Match match in matches)
            {
                if (match.Groups[2].Value == skip)
                {
                    continue;
                }
                var temp = new ArrayList();
                temp.Add(DateTime.Parse($"{match.Groups[1].Value}/{GetYear()}"));
                temp.Add(match.Groups[2].Value);
                temp.Add(double.Parse(match.Groups[3].Value));
                ExpenseList.Add(temp);
            }
            return ExpenseList;
        }
    }
}
