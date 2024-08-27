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
            //if (!worker.CheckDuplicatePdf())
            //{
            //    Database db = new Database();
            //    db.CreateTable(worker.CreateExpenseList());
            //    db.CloseDatabase();
            //}
            Spreadsheet ss = new Spreadsheet(worker.CreateExpenseList());
            ss.AddToExcel();
            //Console.WriteLine("Enter name:");
            //string? aqui = Console.ReadLine();
            //Console.WriteLine($"hola, {aqui}");
        }

        private PdfWorker CreateWorker()
        {
            string PdfNameStorage = @"D:/afterGrad/c#/Adelisa/HerBudget/pdfs/idStore.txt";
            string ReBank = "A\\.pdf|C\\.pdf";
            Match m = Regex.Match(this.PathPdf, ReBank);
            PdfWorker? worker = null;
            switch (m.Value)
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
