using System.Collections.Generic;

namespace ExcelReader
{
    public class Config
    {
        public string BookName { get; }
        public string SheetName { get; }
        public int? StartRow { get; }
        public int? EndRow { get; }
        public Dictionary<string, string> PropertyColumns { get; }

        public Config()
        {
            BookName = "SomeData.xlsx";
            SheetName = "Sheet1";
            StartRow = null;
            EndRow = null;

            PropertyColumns = new Dictionary<string, string>();
            PropertyColumns.Add("Data:Company", "C");
            PropertyColumns.Add("Data:Address", "D");
            PropertyColumns.Add("Data:City",    "E");
            PropertyColumns.Add("Data:Counts:0:Name",  "F");
            PropertyColumns.Add("Data:Counts:0:Value", "G");
            PropertyColumns.Add("Data:Counts:1:Name",  "H");
            PropertyColumns.Add("Data:Counts:1:Value", "I");
        }
    }
}
