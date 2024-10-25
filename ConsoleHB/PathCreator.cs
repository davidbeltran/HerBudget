/*
 * Author: David Beltran
 */

using System;
using System.IO;

namespace ConsoleHB
{
    /// <summary>
    /// Class to allow local paths be found from any PC
    /// </summary>
    public class PathCreator
    {
        public string NewDirectory { get; set; }
        public string NewFile { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="NewDirectory">directory name</param>
        /// <param name="NewFile">file name</param>
        public PathCreator(string NewDirectory, string NewFile)
        {
            this.NewDirectory = $"\\{NewDirectory}";
            this.NewFile = $"\\{NewFile}";
        }

        /// <summary>
        /// Creates directory if non-existant
        /// </summary>
        /// <returns>string with new full path to directory</returns>
        private string MakeDirectory()
        {
            string path = Directory.GetParent(Environment.CurrentDirectory)!.Parent!.FullName + this.NewDirectory;
            string otro = Directory.GetParent(Environment.CurrentDirectory)!.Parent!.FullName;
            Console.WriteLine($"HERE: {otro}");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        /// <summary>
        /// Adds the file name to its appropriate directory
        /// </summary>
        /// <returns>string with complete path including file's name</returns>
        public string MakeFile()
        {
            return MakeDirectory() + this.NewFile;
        }
    }
}
