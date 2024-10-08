﻿using System;
using System.Reflection;

using OfficeOpenXml;

namespace OLT.Extensions.EPPlus.Tests
{
    public class TestBase : IDisposable
    {
        protected readonly ExcelPackage ExcelPackage1;
        protected readonly ExcelPackage ExcelPackage2;
        private const string ResourceName1 = "OLT.Extensions.EPPlus.Tests.Resources.testsheets1.xlsx";
        private const string ResourceName2 = "OLT.Extensions.EPPlus.Tests.Resources.testsheets2.xlsx";

        protected readonly ExcelPackage ImportTestExcelPackage;
        private const string ImportTestResourceName = "OLT.Extensions.EPPlus.Tests.Resources.ImportTest.xlsx";

        protected TestBase()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage1 = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(ResourceName1));
            ExcelPackage2 = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(ResourceName2));
            ImportTestExcelPackage = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(ImportTestResourceName));
        }

        public void Dispose()
        {
            ExcelPackage1.Dispose();
            ExcelPackage2.Dispose();
            ImportTestExcelPackage.Dispose();
        }

        protected string GetRandomName() => new Random(Guid.NewGuid().GetHashCode()).Next().ToString();
    }
}
