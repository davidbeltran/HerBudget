/*
 * Author: David Beltran
 */

namespace ConsoleHB
{
    internal class Program
    {
        static void Main(string[] args)
        {
            PathCreator pc = new PathCreator("pdfs", "MarApr24C.pdf");
            string pathPdf = pc.MakeFile();

            Statement stmt = new Statement(pathPdf);
            stmt.SendToDatabase();
        }
    }
}
/*
 * To ensure program can be ran on different project:
 * - Navigate to Project > Add Project References... > Browse
 * - Select both OFFICE.DLL and Microsoft.Office.Interop.Excel.dll
 * - Make sure to clean and build solution.
 * - Should work properly now.
*/