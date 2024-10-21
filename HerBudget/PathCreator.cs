namespace HerBudget
{
    public class PathCreator
    {
        public string NewDirectory { get; set; }
        public string NewFile { get; set; }

        public PathCreator(string NewDirectory, string NewFile)
        {
            this.NewDirectory = $"\\{NewDirectory}";
            this.NewFile = $"\\{NewFile}";
        }

        private string MakeDirectory()
        {
            string path = Directory.GetParent(Environment.CurrentDirectory)!.Parent!.FullName + this.NewDirectory;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        public string MakeFile()
        {
            return MakeDirectory() + this.NewFile;
        }
    }
}
