/*
 * Author: David Beltran
 */

using System;
using System.Text.RegularExpressions;

namespace ConsoleHB
{
    /// <summary>
    /// Used to store individual expense data
    /// </summary>
    public class Expense
    {
        public DateTime Date {  get; set; }
        public string Detail { get; set; }
        public double Amount { get; set; }
        public CategoryType? Category { get; set; }
        public SubCategoryType? SubCategory { get; set; }
        public string Month { get; set; }
        public string Year { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="date">date of expense</param>
        /// <param name="detail">describes where purchase was made</param>
        /// <param name="Amount">usd cost of purchase</param>
        public Expense(DateTime date, string detail, double Amount)
        {
            this.Date = date;
            this.Detail = detail;
            this.Amount = Amount;
            Categorize();
            this.Month = date.ToString("MMM"); // Used to name excel worksheet
            this.Year = date.Year.ToString();
        }

        /// <summary>
        /// First step to categorize between a bill or expense
        /// </summary>
        private void Categorize()
        {
            string ReCat = "CENTRE CLUB|SPECTRUM|UNITED FIN CAS INS|SO CAL EDISON|SO CAL GAS|DISNEY|DEPT EDUCATION STUDENT LN|" +
                "ATT PAYMENT|HULU|PEACOCK|SLING|KAISER|DELTACARE";

            if (Regex.IsMatch(this.Detail, ReCat))
            {
                this.Category = CategoryType.BILL;
                SubCategorize(ReCat);
            }
            else { this.Category = CategoryType.EXPENSE; }
        }

        /// <summary>
        /// loads subcategory of Expense object to allow organization on Excel sheet
        /// </summary>
        /// <param name="reCat">RegEx pattern to find subcategory description</param>
        private void SubCategorize(string reCat)
        {
            Match m = Regex.Match(this.Detail, reCat);
            switch (m.Value)
            {
                case "CENTRE CLUB":
                    this.SubCategory = SubCategoryType.RENT; break;
                case "SPECTRUM":
                    this.SubCategory = SubCategoryType.INTERNET; break;
                case "UNITED FIN CAS INS":
                    this.SubCategory = SubCategoryType.CAR_INSURANCE; break;
                case "SO CAL EDISON":
                    this.SubCategory = SubCategoryType.ELECTRIC; break;
                case "SO CAL GAS":
                    this.SubCategory = SubCategoryType.GAS_HOME; break;
                case "ATT PAYMENT":
                    this.SubCategory = SubCategoryType.PHONES; break;
                case "HULU" or "PEACOCK" or "SLING" or "DISNEY":
                    this.SubCategory = SubCategoryType.TV; break;
                case "KAISER":
                    this.SubCategory = SubCategoryType.HEALTHCARE; break;
                case "DELTACARE":
                    this.SubCategory = SubCategoryType.DENTAL; break;
                case "DEPT EDUCATION STUDENT LN":
                    this.SubCategory = SubCategoryType.STUD_LN; break;
                default:
                    this.SubCategory = null; break;
            }
        }
    }
}
