﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OLT.Extensions.EPPlus.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace OLT.Extensions.EPPlus.Tests
{
    public class ExcelTemplateGeneratorTests
    {
        [Fact]
        public void Should_generate_an_Excel_package_from_given_ExcelExportable_class_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type wrongCarsType = executingAssembly.GetTypesMarkedAsExcelWorksheet().First(x => x.Name == "WrongCars");
            var worksheetIndex = 0;


            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage excelPackage1 = executingAssembly.GenerateExcelPackage(wrongCarsType.Name);

            Action act = () => executingAssembly.GenerateExcelPackage("sadas");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            excelPackage1.Should().NotBe(null);
            excelPackage1.GetWorksheet(worksheetIndex).GetColumns(1).Count().Should().BeGreaterThan(0);

            act.Should().Throw<InvalidOperationException>().WithMessage("The 'sadas' type could not found in the assembly.");
        }

        [Fact]
        public void Should_generate_an_worksheet_from_given_ExcelExportable_class_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type wrongCarsType = executingAssembly.GetTypesMarkedAsExcelWorksheet().First(x => x.Name == "WrongCars");
            KeyValuePair<string, string> defaultMapType = executingAssembly.GetExcelWorksheetNamesOfMarkedTypes().First(x => x.Key == "DefaultMap");

            var excelPackage = new ExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GenerateWorksheet(executingAssembly, wrongCarsType.Name);
            ExcelWorksheet worksheet2 = excelPackage.GenerateWorksheet(executingAssembly, defaultMapType.Key,
                                                                       act => act.SetHorizontalAlignment(ExcelHorizontalAlignment.Right)
                                                                                 .SetFontAsBold());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet1.Should().NotBe(null);
            worksheet1.Name.Should().Be("Wrong Cars");
            worksheet1.GetColumns(1).Count().Should().BeGreaterThan(0);

            worksheet2.Should().NotBe(null);
            worksheet2.Name.Should().Be("DefaultMap");
            worksheet2.GetColumns(1).Count().Should().BeGreaterThan(0);
            worksheet2.Cells[1, 1].Style.HorizontalAlignment.Should().Be(ExcelHorizontalAlignment.Right);
            worksheet2.Cells[1, 2].Style.Font.Bold.Should().BeTrue();
        }
    }
}
