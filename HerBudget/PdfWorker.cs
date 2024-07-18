using System.Collections;
using System.Text.RegularExpressions;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig;

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
        /// Finds PDF file name in storage file if it exists
        /// </summary>
        /// <returns>true if file name exists. false if it does not exist</returns>
        protected bool SearchForPdf()
        {
            if (File.Exists(this.FileStorage))
            {
                try
                {
                    using StreamReader sr = new StreamReader(this.FileStorage);
                    string PdfFiles = sr.ReadToEnd();
                    if (Regex.IsMatch(PdfFiles, this.PdfDoc))
                    {
                        return true;
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine("Error reading the idStore.txt file", ex.Message);
                }
            }
            else
            {
                File.CreateText(this.FileStorage).Close();
            }
            return false;
        }

        /// <summary>
        /// Writes pdf file name into storage if file has never been processed
        /// </summary>
        /// <returns>returns false if pdf file has never been processed</returns>
        public bool CheckDuplicatePdf()
        {
            if (!SearchForPdf())
            {
                try
                {
                    using StreamWriter sw = new StreamWriter(this.FileStorage, true);
                    sw.WriteLine(this.PdfDoc);
                }
                catch (IOException ex)
                {
                    Console.WriteLine(ex.Message);
                }
                return false;
            }
            else
            {
                Console.WriteLine($"{this.PdfDoc} has already been processed.");
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

        protected string PreparePdf(string pdfDoc)
        {
            string PageText = "";
            try
            {
                using (PdfDocument doc = PdfDocument.Open(pdfDoc))
                {
                    PageText = ContentOrderTextExtractor.GetText(doc.GetPages());//.GetPage(3)
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException("**PDF file not included in folder. Copy from Python project**", ex.Message);
            }
            return PageText;
        }

        public abstract ArrayList CreateExpenseList();
    }
}
