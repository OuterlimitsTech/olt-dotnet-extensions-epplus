using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OLT.Extensions.EPPlus
{
    public static class ExcelCsvConverter
    {

        /// <summary>
        /// Converts worksheet to CSV
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns><see cref="CsvWorksheet"/></returns>
        public static CsvWorksheet ToCsv(this ExcelWorksheet worksheet)
        {
            var maxColumnNumber = worksheet.Dimension.End.Column;
            var currentRow = new List<string>(maxColumnNumber);
            var totalRowCount = worksheet.Dimension.End.Row;
            var currentRowNum = 1;


            var memory = new MemoryStream();

            using (var writer = new StreamWriter(memory, Encoding.ASCII))
            {
                while (currentRowNum <= totalRowCount)
                {
                    BuildRow(worksheet, currentRow, currentRowNum, maxColumnNumber);
                    WriteRecordToFile(currentRow, writer, currentRowNum, totalRowCount);
                    currentRow.Clear();
                    currentRowNum++;
                }
            }

            return new CsvWorksheet(worksheet.Name, memory.ToArray());

        }

        /// <summary>
        /// Converts all Worksheets in Workbook to a CSV byte array
        /// </summary>
        /// <param name="package"></param>
        /// <returns>Returns List of <see cref="CsvWorksheet"/></returns>
        public static List<CsvWorksheet> ToCsv(this ExcelPackage package)
        {
            var list = new List<CsvWorksheet>();

            package.Workbook.Worksheets
                .ToList()
                .ForEach(worksheet =>
                {
                    list.Add(worksheet.ToCsv());
                });

            return list;
        }


        private static void WriteRecordToFile(List<string> cellValues, StreamWriter sw, int rowNumber, int totalRowCount)
        {
            var commaDelimitedRecord = string.Join(",", cellValues);

            if (rowNumber == totalRowCount)
            {
                sw.Write(commaDelimitedRecord);
            }
            else
            {
                sw.WriteLine(commaDelimitedRecord);
            }
        }

        private static void BuildRow(ExcelWorksheet worksheet, List<string> currentRow, int currentRowNum, int maxColumnNumber)
        {
            for (int i = 1; i <= maxColumnNumber; i++)
            {
                AddCellValue(GetCellText(worksheet.Cells[currentRowNum, i]), currentRow);
            }
        }

        /// <summary>
        /// Can't use .Text: http://epplus.codeplex.com/discussions/349696
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string? GetCellText(this ExcelRangeBase cell)
        {
            return cell?.Value?.ToString();
        }

        private static void AddCellValue(string? value, List<string> record)
        {
            record.Add(string.Format("{0}{1}{0}", '"', value ?? string.Empty));
        }
    }
}