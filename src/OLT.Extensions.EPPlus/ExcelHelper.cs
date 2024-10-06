using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OLT.Extensions.EPPlus
{


    public static class ExcelHelper
    {
        /// <summary>
        /// Calculates the Excel Column Index for a giving Letter
        /// https://www.vishalon.net/blog/excel-column-letter-to-number-quick-reference
        /// </summary>
        /// <param name="columnLetter">
        /// Must only be 1-2 alpha characters
        /// </param>
        /// <returns>index or throws <seealso cref="ArgumentOutOfRangeException"/> or <seealso cref="ArgumentNullException"/></returns>
        public static int ColumnLetterToColumnIndex(string columnLetter)
        {
            if (columnLetter == null)
            {
                throw new ArgumentNullException(nameof(columnLetter));
            }

            var isValid = columnLetter.Length > 0 && columnLetter.Length <= 2 && columnLetter.All(char.IsLetter);
            if (!isValid)
            {
                throw new ArgumentOutOfRangeException(nameof(columnLetter), "Must only be 1-2 alpha characters");
            }

            columnLetter = columnLetter.ToUpper();
            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += columnLetter[i] - 'A' + 1;
            }
            return sum;
        }

        /// <summary>
        /// Calculates the Excel Column Index for a giving Letter
        /// https://www.vishalon.net/blog/excel-column-letter-to-number-quick-reference
        /// </summary>
        /// <param name="colIndex">
        /// Must be between 1 and 702
        /// </param>
        /// <returns>string or throws <seealso cref="ArgumentOutOfRangeException"/></returns>
        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            var isValid = colIndex > 0 && colIndex <= 702;
            if (!isValid)
            {
                throw new ArgumentOutOfRangeException(nameof(colIndex), "Must be between 1 and 702");
            }

            int div = colIndex;
            string colLetter = string.Empty;

            while (div > 0)
            {
                var mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }



    }
}