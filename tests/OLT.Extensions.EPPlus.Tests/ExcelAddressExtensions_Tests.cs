﻿using FluentAssertions;

using OfficeOpenXml;

using Xunit;

namespace OLT.Extensions.EPPlus.Tests
{
    public class ExcelAddressExtensionsTests : TestBase
    {
        [Fact]
        public void Should_given_address_range_be_empty_with_header_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddressBase address = ExcelPackage1.GetWorksheet("TEST7").GetValuedDimension();
            ExcelAddressBase address2 = ExcelPackage1.Workbook.Worksheets.Add(GetRandomName()).ChangeCellValue(1, 1,"").GetValuedDimension();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result = address.IsEmptyRange(true);
            bool result2 = address2.IsEmptyRange(true);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result.Should().BeTrue();
            result2.Should().BeTrue();
        }

        [Fact]
        public void Should_given_address_range_not_be_empty_without_header_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddressBase address = ExcelPackage1.GetWorksheet("TEST7").GetValuedDimension();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result = address.IsEmptyRange();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result.Should().BeFalse();
        }
    }
}
