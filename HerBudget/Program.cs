/*
 * Author: David Beltran
 */

namespace HerBudget
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string pathPdf = "D:/afterGrad/c#/Adelisa/HerBudget/pdfs/MarApr24C.pdf";

            Statement stmt = new Statement(pathPdf);
            stmt.SendToDatabase();

            WebStarter ws = new WebStarter(args);
            ws.ShowWeb();
        }
    }
}
