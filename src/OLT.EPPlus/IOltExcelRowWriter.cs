﻿using OfficeOpenXml;
using System;

namespace OLT.EPPlus
{
    [Obsolete("Package No Longer Maintained -> Use OLT.Extensions.EPPlus")]
    public interface IOltExcelRowWriter
    {
        /// <summary>
        /// Writes a row to the worksheet at the given row index and returns the next row idx
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row">Current Row</param>
        /// <returns>Row Number assigned</returns>
        int Write(ExcelWorksheet worksheet, int row);
    }
}