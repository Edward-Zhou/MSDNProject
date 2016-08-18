using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using WinForm.ClassLibrary;

namespace WinForm.Forms
{
    public partial class Classfrm : Form
    {
        public Classfrm()
        {
            InitializeComponent();
        }

        private void operationbtn_Click(object sender, EventArgs e)
        {
            Methods.operation();
        }

        private void getFiles_Click(object sender, EventArgs e)
        {
            List<string> files = new List<string>();
            files = Methods.getFile(@"D:\work\Palladium\Thread Practise\201507");
            foreach (var item in files)
            {
                label1.Text += item + "\r\n";
            }
        }

        private void convertFrombtn_Click(object sender, EventArgs e)
        {
            Object test = "Hello";
            // string t = StringConverter.ConvertFrom(,,);
        }

        private void Linqbtn_Click(object sender, EventArgs e)
        {
            Records records1 = new Records() { A = 0, B = 1, C = 4 };
            Records records2 = new Records() { A = 0, B = 2, C = 1 };
            Records records3 = new Records() { A = 1, B = 2, C = 2 };
            Records records4 = new Records() { A = 2, B = 0, C = 3 };
            List<Records> records = new List<Records>();
            records.Add(records1);
            records.Add(records2);
            records.Add(records3);
            records.Add(records4);
            var t1 = records.GroupBy(r => r.A)
                .SelectMany(rr => rr.Select(t => new { A = t.A })).ToList();
            var t2 = records.GroupBy(r => r.A)
    .Select(rr => rr.Select(t => new { A = t.A })).ToList();
            var finalresult = records.GroupBy(r => r.A)
                .SelectMany(rr => rr.Select(
                        record => new
                        {
                            A = record.A,
                            B = rr.Sum(c => c.C),
                            C = (float)record.C / (float)rr.Sum(c => c.C)
                        }
                    )).ToList();
        }
        class Records
        {
            public int A { get; set; }
            public int B { get; set; }
            public int C { get; set; }
        }

        private void ConvertFucbtn_Click(object sender, EventArgs e)
        {
            System.Linq.Expressions.Expression<Func<DepartmentViewModel, bool>> expr = i => i.name == "DepartmentViewModel";


            Func<DepartmentViewModel, bool> some_function = expr.Compile();
            Func<Department, bool> converted = d => some_function(
                            new DepartmentViewModel
                            {
                                name = d.name
                            }
                            );
            expr.Compile();
            DepartmentViewModel dtv = new DepartmentViewModel();
            dtv.name = "DepartmentViewModel";
            bool b1 = some_function(dtv);

            Department dt = new Department();
            dt.name = "Department";
            bool b2 = converted(dt);
            //bool t1 = test(dtv);
            //bool t2 = test(dt);

            Expression<Func<DepartmentViewModel, bool>> srcLambda = i => i.name == "DepartmentViewModel";
            //Expression<Func<Department, bool>> destLambda = ConvertTo<Department, DepartmentViewModel>(srcLambda);
        }



        public static Expression<Func<TDest, bool>> ConvertTo<TSrc, TDest>(Expression<Func<TSrc, bool>> srcExp)
        {
            ParameterExpression destPE = Expression.Parameter(typeof(TDest));

            ExpressionConverter ec = new ExpressionConverter(typeof(TSrc), destPE);
            Expression body = ec.Visit(srcExp.Body);
            return Expression.Lambda<Func<TDest, bool>>(body, destPE);
        }
        public class ExpressionConverter : ExpressionVisitor
        {

            private Type srcType;
            private ParameterExpression destParameter;

            public ExpressionConverter(Type src, ParameterExpression dest)
            {
                this.srcType = src;
                this.destParameter = dest;
            }

            protected override Expression
               VisitParameter(ParameterExpression node)
            {
                if (node.Type == srcType)
                    return this.destParameter;
                return base.VisitParameter(node);
            }
        }

        private IEnumerable<string> extra(string[] a, IEnumerable<string> b)
        {

            Array.Sort(a); // use a copy of the original array if it's not acceptable to sort it.
            Array.Sort(b.ToArray()); // idem
            int i = 0;
            IEnumerable<string> customDifference = a
                .Where(sl =>
                {
                    int ix = Array.FindIndex(b.ToArray(), i, sr => sr == sl);
                    if (ix >= 0) i = ix + 1;
                    return ix == -1 ? true : false;
                });
            return customDifference;
        }

        private void getDistinct_Click(object sender, EventArgs e)
        {
            string[] a = { "a", "a", "a", "c", "c", "e" };
            string[] b = { "a", "b", "c", "a" };

            List<string> la = a.ToList();
            List<string> lb = b.ToList();
            List<string> result = new List<string>();
            getValue(ref la, ref lb, ref result);
        }

        public void getValue(ref List<string> a, ref List<string> b, ref List<string> result)
        {
            if (a.Count() > 0)
            {
                int index = b.IndexOf(a[0]);
                if (index > -1)
                {
                    a.RemoveAt(0);
                    b.RemoveAt(index);
                    getValue(ref a, ref b, ref result);
                }
                else
                {
                    result.Add(a[0]);
                    a.RemoveAt(0);
                    getValue(ref a, ref b, ref result);
                }
            }
        }

        private void AutoProperty_Click(object sender, EventArgs e)
        {
            LogEventInfoData le = new LogEventInfoData("test", "hello");

        }

        private void SubStringbtn_Click(object sender, EventArgs e)
        {
            String replacement = "";
            String sentence = "TAG123_Sample";
            String pattern = @"TAG[1-9]{1,3}_";
            String sentence1 = "TAG_Sample";
            String pattern1 = @"TAG_";
            Regex r = new Regex(pattern);

            Regex r1 = new Regex(pattern1);
            String res = r.Replace(sentence, replacement);
            String res1 = r1.Replace(sentence1, replacement);
            Console.WriteLine(res);
            Console.ReadLine();

        }
        static void Foo<T>(T arg)
        {
            Console.WriteLine("T arg got called! " + arg);
        }
        static void Foo<T>(MyClass<T> arg)
        {
            Console.WriteLine("MyClass<T> arg got called! " + arg);
        }
        static void Foo<T>(List<T> arg)
        {
            Console.WriteLine("List<T> arg got called! " + arg);
        }

        private void CallNull_Click(object sender, EventArgs e)
        {
            Console.WriteLine();
            Foo<object>(null);
            Foo<object>(null as object);
            Foo<object>(default(object));
        }

        private void XmlReadbtn_Click(object sender, EventArgs e)
        {
            XmlTextReader reader = new XmlTextReader(@"D:\Edward\Project\MSDNProject\MSDNProject\WinForm\XMLFile2.xml");
            StringBuilder output = new StringBuilder();

            XmlWriterSettings ws = new XmlWriterSettings();
            ws.Indent = true;
            using (XmlWriter writer = XmlWriter.Create(output, ws))
            {

                // Parse the file and display each of the nodes.
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element: // The node is an element.
                            Console.Write("<" + reader.Name);
                            Console.WriteLine(">");
                            break;
                        case XmlNodeType.Text: //Display the text in each element.
                            Console.WriteLine(reader.Value);
                            break;
                        case XmlNodeType.EndElement: //Display the end of the element.
                            Console.Write("</" + reader.Name);
                            Console.WriteLine(">");
                            break;
                    }
                }

              
            }
            DataSet ds = new DataSet();
            ds.ReadXml(@"D:\Edward\Project\MSDNProject\MSDNProject\WinForm\XMLFile2.xml");
            DataTable dt = new DataTable();
            dt = ds.Tables[0];
            DataColumn dc = new DataColumn("Id");
            dc.DataType=typeof(Int32);
            dt.Columns.Add(dc);
            foreach (DataRow dr in dt.Rows)
            {
                dr["Id"] = Int32.Parse(dr["OrderID"].ToString());
            }
            dt.DefaultView.Sort = "Id asc";
            this.dataGridView1.DataSource = dt;
            //private void button5_Click(object sender, EventArgs e)
            //{

            //    string[] a = { "a", "a", "a", "c", "c", "e" };
            //    string[] b = { "a", "b", "c", "a" };

            //    this.extra(a, b);

            //} 
        }

        private void TaskRunbtn_Click(object sender, EventArgs e)
        {
            //Task<int> taskOriginal = Task.Run(() =>
            //{
            //    Console.WriteLine("Returning 42");
            //    return 42;
            //});
            //var completedTask = taskOriginal.ContinueWith((i) =>
            //{
            //    Console.WriteLine("OnlyOnRanToCompletion");
            //}, TaskContinuationOptions.OnlyOnRanToCompletion);
            //completedTask.Wait();
            Task<int> taskOriginal = Task.Run(() =>
            {
                Console.WriteLine("Returning 42");
                return 42;
            });
            taskOriginal.ContinueWith((i) =>
            {
                Console.WriteLine("OnlyOnRanToCompletion");
            }, TaskContinuationOptions.OnlyOnRanToCompletion);

            taskOriginal.Wait();


        }

        private void RegexBtn_Click(object sender, EventArgs e)
        {
            string regex = "2015/10/10";
            string regex1 = "2015-10-10";
            Regex r = new Regex(@"(^[0-9]{4,4}/?[0-1][0-9]/?[0-3][0-9]$)");
            bool b = r.IsMatch(regex1);
            Console.WriteLine(b.ToString());
        }

        private void ExcelData_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtExcel = new System.Data.DataTable();
            dtExcel.TableName = "Mail";
            string SourceCon = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\v-tazho\Desktop\Test.xlsx;Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
            OleDbConnection con = new OleDbConnection(SourceCon);
            string query = "Select * from [Mail$]";
            OleDbDataAdapter data = new OleDbDataAdapter(query, con);
            data.Fill(dtExcel);
            dataGridView1.DataSource = dtExcel;
        }
    }
    class MyClass<T> : List<T>
    {
    }
    class Department
    {
        public string name;
    }
    class DepartmentViewModel
    {
        public string name;
    }

}
