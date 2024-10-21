namespace HerBudget
{
    public class PathCreator
    {
        public string NewDirectory { get; set; }
        public string NewFile { get; set; }

        public PathCreator(string NewDirectory, string NewFile)
        {
            this.NewDirectory = NewDirectory;
            this.NewFile = NewFile;
        }
    }
}
