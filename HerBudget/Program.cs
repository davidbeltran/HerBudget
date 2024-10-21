/*
 * Author: David Beltran
 */

namespace HerBudget
{
    internal class Program
    {
        static void Main(string[] args)
        {
            PathCreator pc = new PathCreator("HerBudget\\pdfs", "MarApr24C.pdf");
            string pathPdf = pc.MakeFile();

            Statement stmt = new Statement(pathPdf);
            stmt.SendToDatabase();

            WebStarter ws = new WebStarter(args);
            ws.ShowWeb();
        }
    }
}
