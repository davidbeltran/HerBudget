using System.Collections;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class AllyPdfWorker : PdfWorker
    {
        public AllyPdfWorker(string fileStorage, string pdfDoc) : base(fileStorage, pdfDoc)
        {
            this.ReDetail = "(?:((?:0[1-9]|1[1,2])/(?:0[1-9]|[1-2][0-9]|3[0-1])/(?:\\d{4})) " +
            "(Check Card Purchase|ACH Withdrawal|Direct Deposit|Interest Paid|" +
            "WEB Funds Transfer|NOW Withdrawal|NOW Deposit|eCheck Deposit)\\s" +
            "(\\n.*\\s)?(?:.*\\s)*?\\$(.+\\.\\d{2}) -\\$(.+\\.\\d{2}) (?:-?\\$.+\\.\\d{2}[\\s|A]))";
        }
        public override ArrayList CreateExpenseList()
        {
            string pdfText = PreparePdf(this.PdfDoc);
            MatchCollection matches = Regex.Matches(pdfText, this.ReDetail);
            ArrayList Expenses = new ArrayList();

            foreach (Match match in matches)
            {
                DateTime date = DateTime.Parse(match.Groups[1].Value.Trim());
                string detail1 = match.Groups[2].ToString().Trim().ToUpper();
                string detail2 = match.Groups[3].ToString().Trim().ToUpper();
                double amount1 = double.Parse(match.Groups[4].Value.Trim());
                double amount2 = double.Parse((match.Groups[5].Value.Trim()));

                Expense exp = new Expense(date, detail2, amount2);
                Expenses.Add(exp);
            }
            foreach (Expense exp in  Expenses)
            {
                //if (exp.Category.Equals(CategoryType.BILL))
                //{
                //    Console.WriteLine($"detail: {exp.Detail} || category: {exp.Category} || sub: {exp.SubCategory}\n");
                //}
                Console.WriteLine($"detail: {exp.Detail} || category: {exp.Category} || sub: {exp.SubCategory}\n");
            }
            //foreach (ArrayList exp in ExpenseList)
            //{
            //    Console.WriteLine($"Date: {exp[0]} || Detail1:{exp[1]} ||Detail2:{exp[2]}" +
            //        $"|| Amount1: {exp[3]} || Amount2: {exp[4]}\n");
            //}
            //StreamWriter sw = new StreamWriter("D:/afterGrad/c#/Adelisa/HerBudget/pdfs/test.txt");
            //sw.WriteLine(pdfText);
            return Expenses;
        }
    }
}