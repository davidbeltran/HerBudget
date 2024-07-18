﻿/*
 * Author: David Beltran
 */

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
            ChasePdfWorker cpw = new ChasePdfWorker(PdfNameStorage, this.PathPdf);
            AllyPdfWorker apw = new AllyPdfWorker(PdfNameStorage, this.PathPdf);
            apw.CreateExpenseList();
            //if (!cpw.CheckDuplicatePdf())
            //{
            //    Database db = new Database();
            //    db.CreateTable(cpw.CreateExpenseList());
            //    db.CloseDatabase();
            //}
        }
    }
}

