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
    /// Manages expense data
    /// </summary>
    public class Statement
    {
        public string PathPdf { get; set; }

        /// <summary>
        /// constructor with pdf file name
        /// </summary>
        /// <param name="pathPdf"></param>
        public Statement(string pathPdf)
        {
            this.PathPdf = pathPdf;
        }

        /// <summary>
        /// Uses Database class to fill MySQL table with pdf expenses
        /// </summary>
        public void SendToDatabase()
        {
            string PdfNameStorage = @"D:/afterGrad/c#/Adelisa/HerBudget/idStore.txt";
            WorkerPdf wp = new WorkerPdf(PdfNameStorage, this.PathPdf);
            if (!wp.CheckDuplicatePdf())
            {
                Database db = new Database();
                db.CreateTable(wp.CreateExpenseList());
                db.CloseDatabase();
            }
        }
    }
}

