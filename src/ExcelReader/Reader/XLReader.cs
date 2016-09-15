using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelReader
{
    public class XLReader
    {
        private XLWorkbook _book;
        private IXLWorksheet _sheet;
        private int? _startRow;
        private int? _endRow;
        private Dictionary<string, string> _propertyColumns;

        public XLReader(Config config)
        {
            _book = new XLWorkbook(config.BookName);
            _sheet = _book.Worksheet(config.SheetName);
            _startRow = config.StartRow;
            _endRow = config.EndRow;
            _propertyColumns = config.PropertyColumns;
        }


        public void ExtractData()
        {
            var firstRowUsed = _startRow != null ? _sheet.Row(_startRow.Value) : _sheet.FirstRowUsed();
            var lastRowUsed = _endRow != null ? _sheet.Row(_endRow.Value) : _sheet.LastRowUsed();
                                 
            var dataReadingRow = firstRowUsed.Row(1, lastRowUsed.RowNumber()); // считываемая строка

            int fullDataRowsCount = 0;        // количество строк с полной информацией
            int halfDataOrEmptyRowsCount = 0; // количество пустых или частично заполненных строк

            // если первая строка - заголовок таблицы, то пропустить ее
            //dataRow = dataRow.RowBelow();

            // чтение строк
            while (dataReadingRow.RowNumber() <= lastRowUsed.RowNumber())
            {
                // проверка наличия данных в целевых ячейках
                var readThisRow = true;
                foreach (var cellNum in _propertyColumns.Values)
                {
                    readThisRow = readThisRow & !dataReadingRow.Cell(cellNum).IsEmpty();
                }

                // читать строку, если заполнены все целевые ячейки
                if (readThisRow)
                {
                    fullDataRowsCount++;

                    var dataItem = new Data
                    {
                        Company = dataReadingRow.Cell(_propertyColumns["Data:Company"]).GetString(),
                        Address = dataReadingRow.Cell(_propertyColumns["Data:City"]).GetString() + ", " +
                                  dataReadingRow.Cell(_propertyColumns["Data:Address"]).GetString(),
                        Counts = new List<Count>
                        {
                            new Count
                            {
                                Name =  dataReadingRow.Cell(_propertyColumns["Data:Counts:0:Name"]).GetString(),
                                Value = dataReadingRow.Cell(_propertyColumns["Data:Counts:0:Value"]).GetString(),
                            },
                            new Count
                            {
                                Name =  dataReadingRow.Cell(_propertyColumns["Data:Counts:1:Name"]).GetString(),
                                Value = dataReadingRow.Cell(_propertyColumns["Data:Counts:1:Value"]).GetString(),
                            },
                        },
                    };

                    // вывод на консоль
                    Console.WriteLine($"Row num: {dataReadingRow.RowNumber()}");
                    Console.WriteLine($"Company: {dataItem.Company}");
                    Console.WriteLine($"Address: {dataItem.Address}");
                    Console.Write($"Counts:  [ ");
                    foreach (var count in dataItem.Counts)
                    {
                        Console.Write($"{{{count.Name}, {count.Value}}}");
                        if (dataItem.Counts.IndexOf(count) < dataItem.Counts.Count - 1)
                            Console.Write(", ");
                    }
                    Console.WriteLine(" ]\n");
                }
                else { halfDataOrEmptyRowsCount++; }

                // чтение следующей строки
                dataReadingRow = dataReadingRow.RowBelow();
            }

            Console.WriteLine("\n=== SUMMARY ===\n");
            Console.WriteLine($"Индекс первой строки:     {firstRowUsed.RowNumber()}");
            Console.WriteLine($"Индекс последней строки:  {lastRowUsed.RowNumber()}");
            Console.WriteLine($"Количество полных строк:  {fullDataRowsCount}");
            Console.WriteLine($"Количество пустых строк:  {halfDataOrEmptyRowsCount}");
        }


        public void Preview()
        {
            const int columnsMaxCount = 50;
            const int rowsMaxCount = 50;

            var firstColumn = _sheet.FirstColumn();
            var lastColumn = _sheet.LastColumnUsed().ColumnNumber() < columnsMaxCount ?
                _sheet.LastColumnUsed() :
                _sheet.Column(columnsMaxCount);

            var firstRow = _sheet.FirstRow();
            var lastRow = _sheet.LastRowUsed().RowNumber() < rowsMaxCount ?
                _sheet.LastRowUsed() :
                _sheet.Row(rowsMaxCount);

            var previewData = new string[lastColumn.ColumnNumber(), lastRow.RowNumber()];

            for (var row = firstRow.RowNumber(); row <= lastRow.RowNumber(); row++)
            {
                for (int col = firstColumn.ColumnNumber(); col <= lastColumn.ColumnNumber(); col++)
                {
                    previewData[col-1, row-1] = _sheet.Row(row).Cell(col).GetString();
                    Console.Write(previewData[col-1, row-1] + " \t");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
            
            var rowsUsed = _sheet.LastRowUsed().RowNumber();
            int rowsRoundedCount = rowsUsed < 100 ? 
                lastRow.RowNumber() :
                (int)Math.Round((double)(rowsUsed - rowsUsed / 20) / 10) * 10;

            Console.WriteLine("\n=== SUMMARY ===\n");
            Console.WriteLine($"Примерное кол-во строк: {rowsRoundedCount}");
        }
    }
}
