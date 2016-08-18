using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    public class Class1
    {
        public static List<string> getFile(string path)
        {
            List<string> files = new List<string>();
            List<string> fileNames = new List<string>();
            files = Directory.GetFiles(path).ToList();
            foreach (var item in files)
            {
                fileNames.Add(Path.GetFileName(item));
            }
            return fileNames;
        }
        public static List<string> getFile1(string path)
        {
            List<string> files = new List<string>();
            List<string> fileNames = new List<string>();
            files = Directory.GetFiles(path).ToList();
            foreach (var item in files)
            {
                fileNames.Add(Path.GetFileName(item));
            }
            return fileNames;
        }
        public static List<string> getFile2(string path)
        {
            List<string> files = new List<string>();
            List<string> fileNames = new List<string>();
            files = Directory.GetFiles(path).ToList();
            foreach (var item in files)
            {
                fileNames.Add(Path.GetFileName(item));
            }
            return fileNames;
        }
    }
}
