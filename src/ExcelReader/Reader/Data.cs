using System.Collections.Generic;

namespace ExcelReader
{
    public class Data
    {
        public string Company { get; set; }

        public string Address { get; set; }

        public List<Count> Counts { get; set; }
    }

    public class Count
    {
        public string Name { get; set; }

        public string Value { get; set; }
    }
}
