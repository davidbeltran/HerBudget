/*
 * Author: David Beltran
 */

namespace ConsoleHB
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //PathCreator pc = new PathCreator("pdfs", "MarApr24C.pdf");
            //string pathPdf = pc.MakeFile();

            Statement stmt = new Statement(@"D:\afterGrad\c#\Adelisa\HerBudget\ConsoleHB\bin\pdfs\MarApr24C.pdf");
            stmt.SendToDatabase();
        }
    }
}
