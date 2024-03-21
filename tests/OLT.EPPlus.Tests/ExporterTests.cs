using OfficeOpenXml;
using OLT.EPPlus.Tests.Assets;
using System;
using Xunit;

namespace OLT.EPPlus.Tests
{
    public class ExporterTests
    {

        [Fact]
        [Obsolete]
        public void FileBuilder()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

            var list = PersonModel.FakerList(4);

            var builder = new TestFileBuilder();
            var result = builder.Build(new TestFileBuilderRequest { Data = list });

            Assert.True(result.FileBase64.Length > 0);
        }

    }
}