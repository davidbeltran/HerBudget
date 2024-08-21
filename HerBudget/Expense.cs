using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace HerBudget
{
    public class Expense
    {
        public DateTime Date {  get; set; }
        public string Detail { get; set; }
        public double Amount { get; set; }
        public CategoryType? Category { get; set; }
        public SubCategoryType? SubCategory { get; set; }

        public Expense(DateTime date, string detail, double Amount)
        {
            this.Date = date;
            this.Detail = detail;
            this.Amount = Amount;
            Categorize();
        }
        private void Categorize()
        {
            string ReCat = "CENTRE CLUB|SPECTRUM|UNITED FIN CAS INS|SO CAL EDISON|SO CAL GAS|DISNEY|" +
                "ATT PAYMENT|HULU|PEACOCK|SLING|KAISER|DELTACARE|REQUESTED TRANSFER TO ALLY BANK SAVINGS";

            if (Regex.IsMatch(this.Detail, ReCat))
            {
                this.Category = CategoryType.BILL;
                SubCategorize(ReCat);
            }
            else { this.Category = CategoryType.EXPENSE; }
        }

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
                case "REQUESTED TRANSFER TO ALLY BANK SAVINGS":
                    this.SubCategory = SubCategoryType.SAVINGS; break;
                default:
                    this.SubCategory = null; break;
            }
        }
    }
}
