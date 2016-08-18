using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinForm.ClassLibrary
{
    class Methods
    {
        public static void operation()
        {
            decimal amount1, amount2;
            List<decimal> amounts;
            amount1=Convert.ToDecimal("18.2");
            amount2 = Convert.ToDecimal("25.25");
            amounts=new List<decimal>();
            amounts.Add(amount1);
            amounts.Add(amount2);
            string strDecimal = "72.0798454871";//String.Format("{0:F}", 0.720798454871 * 100);//
            var data=from x in amounts
                     select new { result = Convert.ToDecimal(((amount1 / amount2) * 100).ToString("#0.00")) };
            int index = strDecimal.IndexOf(".");
            if (index == -1 || strDecimal.Length < index + 2 + 1)
            {
                strDecimal = string.Format("{0:F" + 2 + "}");
            }
            else
            {
                int length = index;
                length = index + 2 + 1;

                decimal test =Convert.ToDecimal("72.079207920792079207920792080".Substring(0, length));
            }

        }

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
    }
}
