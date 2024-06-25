/*
 * Author: David Beltran
 */

using System.Collections;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace HerBudget
{
    /// <summary>
    /// Handles pdf formatting and specific data retrieval
    /// </summary>
    public class Statement
    {
        public string PathPdf { get; set; }
        public string PageText { get; set; }
        public string Pattern { get; set; }
        public ArrayList ExpList { get; set; }

        /// <summary>
        /// constructor with given regex pattern
        /// </summary>
        /// <param name="pathPdf"></param>
        /// <param name="pattern">can be replaced with other pattern if needed</param>
        /// <param name="pageText"></param>
        public Statement(string pathPdf, string pattern = "(?:\\n((?:0[1-9]|1[1,2])/(?:0[1-9]|[12][0-9]|3[01]))\\s*(.+)" +
            " ((?:-\\d+\\.\\d{2})|(?:\\d+\\.\\d{2})))", string pageText = "")
        {
            this.PathPdf = pathPdf;
            this.Pattern = pattern;
            this.PageText = pageText;
            this.ExpList = new ArrayList();
        }

        /// <summary>
        /// Prints full list for debugging
        /// </summary>
        /// <param name="expenses">experience list generated on Program.cs</param>
        public void PrintExpList(ArrayList expenses)
        {
            foreach (ArrayList exp in expenses)
            {
                Console.WriteLine($"Date: {exp[0]}, Detail: {exp[1]}, Amount: {exp[2]}");
            }
        }

        /// <summary>
        /// Uses Database class to fill MySQL table with pdf expenses
        /// </summary>
        public void SendToDatabase()
        {
            WorkerPdf wp = new WorkerPdf("idStore.txt", this.PathPdf);
            Database db = new Database();
            db.OpenConnection();
            db.CreateTable(wp.CreateExpenseList());
            db.CloseDatabase();
        }
    }
}


/// Extra code 

/// <summary>
/// This would extract every page from the PDF file
/// </summary>
//foreach(Page page in document.GetPages())
//{
//    var text = ContentOrderTextExtractor.GetText(page);
//    pageText += text;
//}

/// <summary>
/// Takes PDF text and loads to txt file
/// </summary>
//File.WriteAllText(pathTxt, pageText);
