using System;

using OfficeOpenXml;

namespace OLT.Extensions.EPPlus
{
    internal class WorksheetTitleRow
    {
        internal string Title { get; set; } = default!;

        internal Action<ExcelRange>? ConfigureTitle { get; set; }
    }
}