using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    class Class2
    {
        public static List<string> getFile4(string path)
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
