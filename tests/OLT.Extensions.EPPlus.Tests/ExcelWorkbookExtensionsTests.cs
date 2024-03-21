﻿using OLT.Extensions.EPPlus.Style;

using FluentAssertions;

using OfficeOpenXml;

namespace OLT.Extensions.EPPlus.Tests
{
    public class ExcelWorkbookExtensionsTests : TestBase
    {
        [Fact]
        public void Should_check_and_throw_if_named_style_found()
        {

            ExcelWorkbook workbook1 = ExcelPackage2.Workbook;
            ExcelWorkbookStyleExtensions.CreateNamedStyle(workbook1, "Style1", act => act.SetFontAsBold());


            //ExcelWorkbookExtensions.

            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage2.GetWorksheet("RandomOrderedColumns");


            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action act1 = () => ExcelWorkbookStyleExtensions.CreateNamedStyle(workbook1, "Style1", act => act.SetTextVertical());

            
            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            act1.Should().Throw<InvalidOperationException>().WithMessage("The Excel package already has a style with the name of 'Style1'");
            //act2.Should().Throw<ExcelValidationException>().WithMessage("'Barcode' column is duplicated (rowIndex: 1)");
        }
    }
}
