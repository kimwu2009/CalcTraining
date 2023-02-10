using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalcTraining
{
    internal class Formula
    {
        public int Index;
        public int Number1;
        public int Number2;
        public int Number3;
        public string Operator1 = "";
        public string Operator2 = "";
        public string Equal = "";
        public string Result = "        ";

        public int ColumnCount
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(Operator2))
                {
                    return 8;
                }
                else
                {
                    return 6;
                }
            }
        }

        public object[] ToArray()
        {
            if (!string.IsNullOrWhiteSpace(Operator2))
            {
                return new object[] { $"({Index})", Number1, Operator1, Number2, Operator2, Number3, Equal, Result };
            }
            else
            {
                return new object[] { $"({Index})", Number1, Operator1, Number2, Equal, Result };
            }
        }

        override public string ToString()
        {
            if (!string.IsNullOrWhiteSpace(Operator2))
            {
                return $"({Index}) {Number1} {Operator1} {Number2} {Operator2} {Number3} {Equal}";
            }
            else
            {
                return $"({Index}) {Number1} {Operator1} {Number2} {Equal}";
            }
        }
    }
}
