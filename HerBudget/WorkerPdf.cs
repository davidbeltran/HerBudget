/*
 * Author: David Beltran
 */

using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig;
using System.Collections;
using System.Text.RegularExpressions;

namespace HerBudget
{
    /// <summary>
    /// Handles pdf formatting and specific data retrieval
    /// </summary>
    public class WorkerPdf
    {
        private string FileStorage { get; set; }
        private string ReDetail { get; set; }
        private string ReYear { get; set; }
        private string PdfDoc { get; set; }
        private string PageText { get; set; }

        /// <summary>
        /// Constructor with given regex patterns
        /// </summary>
        /// <param name="fileStorage"></param>
        /// <param name="pdfDoc"></param>
        public WorkerPdf(string fileStorage, string pdfDoc)
        {
            this.FileStorage = fileStorage;
            this.PdfDoc = pdfDoc;
            this.ReDetail = "(?:\\n((?:0[1-9]|1[1,2])/(?:0[1-9]|[12][0-9]|3[01]))\\s*(.+)" +
                " ((?:-\\d+\\.\\d{2})|(?:\\d+\\.\\d{2})))";
            this.ReYear = "\\d{2}";
        }

        /// <summary>
        /// Finds PDF file name in storage file if it exists
        /// </summary>
        /// <returns>true if file name exists. false if it does not exist</returns>
        private bool SearchForPdf()
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
        private string GetYear()
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
        /// Retrieves third page of pdf and converts to .txt file
        /// </summary>
        /// <param name="pdfPath">PDF file location</param>
        /// <returns>pdf content in string text</returns>
        private string PreparePdf(string pdfPath)
        {
            try
            {
                using (PdfDocument doc = PdfDocument.Open(pdfPath))
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
        public ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();
            foreach (Match match in matches)
            {
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
