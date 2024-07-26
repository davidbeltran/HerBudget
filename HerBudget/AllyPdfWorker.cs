﻿using System.Collections;
using System.Text.RegularExpressions;
using System;

namespace HerBudget
{
    public class AllyPdfWorker : PdfWorker
    {
        public AllyPdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = "(?:((?:0[1-9]|1[1,2])/(?:0[1-9]|[1-2][0-9]|3[0-1])/(?:\\d{4})) " +
            "(?:Check Card Purchase|ACH Withdrawal|Direct Deposit|WEB Funds Transfer|NOW Withdrawal|NOW Deposit)\\s" +
            "\\n(.*\\s)(?:.*\\s)*?" +
            //"\\$(\\d+\\.\\d{2}) -\\$(\\d+\\.\\d{2}))";
            //"\\$(.+\\.\\d{2}) -\\$(\\d+\\.\\d{2}))";
            "\\$(.+\\.\\d{2}) -\\$(.+\\.\\d{2}) (?:-?\\$.+\\.\\d{2}\\s))";
        }
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList ExpenseList = new ArrayList();
            int count = 0;
            foreach (Match match in matches)
            {
                string detail = match.Groups[2].ToString().Replace("\n", "");
                Console.WriteLine($"Date: {match.Groups[1]} || Detail:{match.Groups[2]} " +
                    $"|| Amount1: {match.Groups[3]} || Amount2: {match.Groups[4]}\n");
                count++;
            }
            Console.WriteLine($"Number: {count}");
            return ExpenseList;
        }
    }
}