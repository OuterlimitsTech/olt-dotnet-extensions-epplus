using System;

using OfficeOpenXml;

namespace OLT.Extensions.EPPlus
{
    internal class WorksheetTitleRow
    {
        internal string Title { get; set; }

        internal Action<ExcelRange> ConfigureTitle { get; set; }
    }
}