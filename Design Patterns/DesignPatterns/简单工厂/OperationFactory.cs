using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Design_Patterns.DesignPatterns.简单工厂
{
    class OperationFactory
    {
        public static Operation createOperate(string operate)
        {
            Operation operation = null;
            switch (operate)
            { 
                case "+":
                    operation = new OperationAdd();
                    break;
                case "-":
                    operation = new OperationSub();
                    break;
                case "*":
                    operation = new OperationMul();
                    break;
                case "/":
                    operation = new OperationDiv();
                    break;
            }
            return operation;
        }
    }
}
