/*
 * Author: David Beltran
 */

using System.Collections;

namespace ConsoleHB
{
    /// <summary>
    /// Class utilized for sorting expenses by comparing both months in each PDF file
    /// </summary>
    public class DateComparer : IComparer
    {
        public int Compare(object? x, object? y)
        {
            Expense? exp1 = x as Expense;
            Expense? exp2 = y as Expense;
            return exp1!.Month.CompareTo(exp2!.Month);
        }
    }
}
