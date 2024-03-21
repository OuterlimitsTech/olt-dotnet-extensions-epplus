using System;

using OfficeOpenXml;

namespace OLT.Extensions.EPPlus
{
    internal class WorksheetColumn<T>
    {
        internal Func<T, object> Map { get; set; } = default!;

        internal string Header { get; set; } = default!;

        internal Action<ExcelColumn>? ConfigureColumn { get; set; }
    }
}
