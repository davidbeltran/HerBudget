using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class Expense
    {
        public DateTime Date {  get; set; }
        public string Detail { get; set; }
        public double Amount { get; set; }
        public CategoryType Category { get; set; }
        public SubCategoryType SubCategory { get; set; }

        public Expense(DateTime date, string detail, double Amount)
        {
            this.Date = date;
            this.Detail = detail;
            this.Amount = Amount;
            Categorize();
        }
        private void Categorize()
        {
            string ReCat = "CENTRE CLUB|SPECTRUM|UNITED FIN CAS INS|SO CAL EDISON|SO CAL GAS|" +
                "ATT PAYMENT|HULU|PEACOCK|SLING|KAISER|DELTACARE|REQUESTED TRANSFER TO ALLY BANK SAVINGS";

            if (Regex.IsMatch(this.Detail, ReCat))
            {
                this.Category = CategoryType.BILL;
            }
            else { this.Category = CategoryType.EXPENSE; }
        }
    }
}
