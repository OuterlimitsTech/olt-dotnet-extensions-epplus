﻿using System;
using System.Collections.Generic;
using System.Reflection;

using OLT.Extensions.EPPlus.Attributes;

using OfficeOpenXml;

using static OLT.Extensions.EPPlus.Helpers.Guard;

namespace OLT.Extensions.EPPlus
{
    public static class ExcelTemplateGenerator
    {
        /// <summary>
        ///     Finds given type name in the assembly, and generates Excel package
        /// </summary>
        /// <param name="executingAssembly"></param>
        /// <param name="typeName"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        public static ExcelPackage GenerateExcelPackage(this Assembly executingAssembly, string typeName, Action<ExcelRange>? action = null)
        {
            var excelPackage = new ExcelPackage();
            excelPackage.GenerateWorksheet(executingAssembly, typeName, action);
            return excelPackage;
        }

        /// <summary>
        ///     Finds given type name in the assembly, and generates Excel worksheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="executingAssembly"></param>
        /// <param name="typeName"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        public static ExcelWorksheet GenerateWorksheet(this ExcelPackage excelPackage, Assembly executingAssembly, string typeName, Action<ExcelRange>? action = null)
        {
            Type? type = executingAssembly.GetExcelWorksheetMarkedTypeByName(typeName);

            ThrowIfConditionMet(type == null, "The '{0}' type could not found in the assembly.", typeName);
            
            List<ExcelTableColumnDetails> headerColumns = type?.GetExcelTableColumnAttributesWithPropertyInfo() ?? new List<ExcelTableColumnDetails>();

            var worksheetName = type?.GetWorksheetName();
            worksheetName = string.IsNullOrEmpty(worksheetName) ? typeName : worksheetName;
            ExcelWorksheet worksheet = excelPackage.AddWorksheet(worksheetName);

            for (var i = 0; i < headerColumns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headerColumns[i].ToString();
                action?.Invoke(worksheet.Cells[1, i + 1]);
            }

            return worksheet;
        }
    }
}
