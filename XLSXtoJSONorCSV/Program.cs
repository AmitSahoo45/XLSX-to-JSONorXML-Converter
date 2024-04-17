using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXtoJSONorCSV.Utility;

namespace XLSXtoJSONorCSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CustomUtilityXLS.ConvertXLXStoDesiredFormat("sample01");

            Console.ReadLine();
        }
    }
}
