using OfficeOpenXml;
using System;

namespace OLT.EPPlus
{
    [Obsolete("Package No Longer Maintained -> Use OLT.Extensions.EPPlus")]
    public interface IOltExcelCellStyle
    {
        /// <summary>
        /// Sets Style to given range of cells
        /// </summary>
        /// <param name="range"></param>
        void ApplyStyle(ExcelRange range);
    }
}