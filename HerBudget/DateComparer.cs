namespace HerBudget
{
    public class DateComparer : IComparer<DateTime>
    {
        public int Compare(DateTime x, DateTime y) 
        {
            return x.CompareTo(y);
        }
    }
}
