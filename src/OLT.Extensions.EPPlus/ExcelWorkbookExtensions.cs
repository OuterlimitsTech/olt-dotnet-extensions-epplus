using System.Linq;

using OfficeOpenXml;

namespace OLT.Extensions.EPPlus
{
    public static class ExcelWorkbookExtensions
    {
        public static ExcelWorksheet? GetWorksheet(this ExcelWorkbook workbook, string worksheetName) => workbook.Worksheets.FirstOrDefault(x => x.Name == worksheetName);

        public static ExcelWorksheet GetWorksheet(this ExcelWorkbook workbook, int worksheetIndex) => workbook.Worksheets[worksheetIndex];
    }
}
