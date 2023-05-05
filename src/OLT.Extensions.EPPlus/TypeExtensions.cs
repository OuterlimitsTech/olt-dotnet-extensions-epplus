﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using OLT.Extensions.EPPlus.Attributes;

using static OLT.Extensions.EPPlus.Helpers.Guard;

namespace OLT.Extensions.EPPlus
{
    internal static class TypeExtensions
    {
        internal static object ChangeType(this object value, Type type)
            => value != null ? Convert.ChangeType(value, type) : null;

        /// <summary>
        ///     Returns PropertyInfo and ExcelTableColumnAttribute pairs of given type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        internal static List<ExcelTableColumnDetails> GetExcelTableColumnAttributesWithPropertyInfo(this Type type)
        {
            List<ExcelTableColumnDetails> columnAttributesWithPropertyInfo = type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                                                                                                 .Select(property =>
                                                                                                 {
                                                                                                     var columnAttribute = property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).FirstOrDefault() as ExcelTableColumnAttribute;
                                                                                                     return new ExcelTableColumnDetails(0, property, columnAttribute);
                                                                                                 })
                                                                                                 .Where(p => p.ColumnAttribute != null)
                                                                                                 .ToList();

            ThrowIfConditionMet(!columnAttributesWithPropertyInfo.Any(), "Given object does not have any '{0}'.", nameof(ExcelTableColumnAttribute));
           
            return columnAttributesWithPropertyInfo;
        }

        /// <summary>
        ///     Returns value of the property name
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        internal static object GetPropertyValue(this object obj, string propertyName) => obj.GetType().GetProperty(propertyName)?.GetValue(obj, null);

        internal static string GetWorksheetName(this Type type)
        {
            Attribute worksheetAttribute = type.GetCustomAttribute(typeof(ExcelWorksheetAttribute), true);
            return (worksheetAttribute as ExcelWorksheetAttribute)?.WorksheetName ?? type.Name; 
        }

        /// <summary>
        ///     Determines whether given type is nullable or not
        /// </summary>
        /// <param name="type">Type to test</param>
        /// <returns>True if type is nullable</returns>
        internal static bool IsNullable(this Type type) => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);

        /// <summary>
        ///     Tests whether given type is numeric or not
        /// </summary>
        /// <param name="type">Type to test</param>
        /// <returns>True if type is numeric</returns>
        internal static bool IsNumeric(this Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }
    }
}
