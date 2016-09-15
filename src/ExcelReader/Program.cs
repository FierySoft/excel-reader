using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelReader
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var config = new Config();

            var reader = new XLReader(config);

            Console.WriteLine("\n==================  EXTRACT DATA  =================\n");

            reader.ExtractData();

            Console.WriteLine("\n----------------------- end -----------------------\n\n");

            Console.WriteLine("\n====================  PREVIEW  ====================\n");

            reader.Preview();

            Console.WriteLine("\n----------------------- end -----------------------\n\n");
        }
    }
}
