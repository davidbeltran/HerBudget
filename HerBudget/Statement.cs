/*
 * Author: David Beltran
 */

using System.Text.RegularExpressions;

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
            PdfWorker worker = CreateWorker();
            if (!worker.CheckDuplicatePdf())
            {
                Database db = new Database();
                db.CreateTable(worker.CreateExpenseList());
                db.CloseDatabase();
                Spreadsheet ss = new Spreadsheet(worker.CreateExpenseList());
                ss.AddToExcel();
            }
        }

        /// <summary>
        /// Utilizes the factory design pattern to instantiate PdfWorker object dependent on which bank statement loaded
        /// </summary>
        /// <returns>PdfWorker object corresponding to bank subclass</returns>
        private PdfWorker CreateWorker()
        {
            PathCreator pc = new PathCreator("storage", "idStore.xml");
            string PdfNameStorage = pc.MakeFile();
            //string PdfNameStorage = @"D:/afterGrad/c#/Adelisa/HerBudget/pdfs/idStore.xml";
            string ReBank = "A\\.pdf|C\\.pdf";
            Match m = Regex.Match(this.PathPdf, ReBank);
            PdfWorker? worker = null;
            switch (m.Value) //More can be added if different banks are used. Each bank will need own subclass
            {
                case "A.pdf":
                    worker = new AllyPdfWorker(PdfNameStorage, this.PathPdf); break;
                case "C.pdf":
                    worker = new ChasePdfWorker(PdfNameStorage, this.PathPdf); break;
            }
            return worker!;
        }
    }
}
