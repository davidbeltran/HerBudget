using System.Runtime.CompilerServices;

namespace HerBudget
{
    public class Expense
    {
        private DateTime Date {  get; set; }
        private string? Detail { get; set; }
        private double Amount { get; set; }
        private CategoryType Category { get; set; }
        private SubCategoryType SubCategory { get; set; }

        public Expense(DateTime date, string detail, double Amount)
        {
            this.Date = date;
            this.Detail = detail;
            this.Amount = Amount;
            Categorize(detail);
        }
        private void Categorize(string detail)
        {
            HashSet<string> BillCat = new HashSet<string>()
            {
                "Centre Club WEB PMTS",
                "SPECTRUM SPECTRUM",
                "UNITED FIN CAS INS PREM~ Future",
                "SO CAL EDISON CO DIRECTPAY~ Future",
                "SO CAL GAS PAID SCGC",
                "ATT Payment",
                "HULU SANTA MONICA, CA, USA",
                "Peacock 6911AA P New York, NY, USA",
                "NBC PEACOCK30 ROCKEFELLER NEW YORK, NY, US",
                "Peacock 612D3 P New York, NY, USA",
                "Sling TV LLC 888-3631777 CO",
                "SLING.COM 888-363-1777",
                "KAISER HPTS 8664734938",
                "DeltaCare PREMIUM~ Future Amount: 14.83 ~ Tran: ACHDW",
                "Requested transfer to ALLY BANK"
            };
            if (BillCat.Contains(detail))
            {
                this.Category = CategoryType.BILL;
            }
            else { this.Category = CategoryType.EXPENSE; }
        }
    }
}
