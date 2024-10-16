using System.Collections;
using System.Text.RegularExpressions;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Xml.Linq;

namespace HerBudget
{
    public abstract class PdfWorker
    {
        protected string FileStorage {  get; set; }
        protected string PdfDoc { get; set; }
        protected string ReYear { get; set; }
        protected string ReDetail { get; set; } = null!;

        protected PdfWorker(string fileStorage, string pdfDoc)
        {
            this.FileStorage = fileStorage;
            this.PdfDoc = pdfDoc;
            this.ReYear = "\\d{2}";
        }

        /// <summary>
        /// Writes pdf file name into XML storage if file has never been processed
        /// </summary>
        /// <returns>returns false if pdf file has never been processed</returns>
        public bool CheckDuplicatePdf()
        {
            if (!File.Exists(this.FileStorage))
            {
                new XDocument(new XElement("PdfFiles")).Save(this.FileStorage);
            }

            XDocument doc = XDocument.Load(this.FileStorage);
            XElement? root = doc.Element("PdfFiles");
            string pattern = @"[^/\\]+\.pdf$";//Regex for finding pdf file name at end of entire path

            if (root != null)
            {
                Match match = Regex.Match(this.PdfDoc, pattern);
                string pdfFile = match.Value;
                bool alreadyExists = root.Elements("PdfFile").Any(e => e.Value == pdfFile);//Linq search
                if (alreadyExists)
                {
                    Console.WriteLine($"{pdfFile} has already been processed.");
                    return true;
                }
                root.Add(new XElement("PdfFile", pdfFile));
                doc.Save(this.FileStorage);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Scrapes pdf file name to find year of expenses
        /// </summary>
        /// <returns>string of year</returns>
        protected string GetYear()
        {
            MatchCollection matches = Regex.Matches(this.PdfDoc, this.ReYear);
            string year = "";
            foreach (Match match in matches)
            {
                year = match.Value;
            }
            return year;
        }

        /// <summary>
        /// Extracts entire PDF content and stores in a string
        /// </summary>
        /// <param name="pdfDoc">full path of pdf file</param>
        /// <returns>pdf content as a string</returns>
        /// <exception cref="ArgumentException">Message if pdf file not found in system</exception>
        protected string PreparePdf(string pdfDoc)
        {
            string PageText = "";
            try
            {
                using (PdfDocument doc = PdfDocument.Open(pdfDoc))
                {
                    int counter = 1;
                    foreach (Page page in doc.GetPages())
                    {
                        PageText += ContentOrderTextExtractor.GetText(doc.GetPage(counter));
                        counter++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException("PDF file not found.", ex.Message);
            }
            return PageText;
        }

        /// <summary>
        /// Each individual bank will have this method catered in inherited class
        /// </summary>
        /// <returns>Needs to return an arraylist of Expense objects</returns>
        public abstract ArrayList CreateExpenseList();
    }
}
