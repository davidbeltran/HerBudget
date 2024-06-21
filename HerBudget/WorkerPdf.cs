namespace HerBudget
{
    public class WorkerPdf
    {
        private string fileStorage { get; set; }
        private string reDetail { get; set; }
        private string reYear { get; set; }
        private string pdfDoc { get; set; }

        public WorkerPdf(string fileStorage, string pdfDoc)
        {
            this.fileStorage = fileStorage;
            this.pdfDoc = pdfDoc;
            this.reDetail = "(?:\\n((?:0[1-9]|1[1,2])/(?:0[1-9]|[12][0-9]|3[01]))\\s*(.+)" +
                " ((?:-\\d+\\.\\d{2})|(?:\\d+\\.\\d{2})))";
            this.reYear = "\\d{2}";
        }
    }
}
