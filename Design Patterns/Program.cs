using Design_Patterns.DesignPatterns.简单工厂;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Design_Patterns
{
    class Program
    {
        /// <summary>
        /// 计算器的代码，实现方法的封装，继承
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Operation oper;
            oper = OperationFactory.createOperate("+");
            oper.NumberA = 10;
            oper.NumberB = 5;
            double result = oper.GetResult();
            Console.WriteLine(result);
            Console.ReadLine();
        }
    }
}
