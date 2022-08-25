﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;

using OLT.Extensions.EPPlus.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace OLT.Extensions.EPPlus.Tests
{
    public class ToExcelExtensionsTests : TestBase
    {
        public ToExcelExtensionsTests() => _personList = new List<Person>
                                                          {
                                                              new Person { FirstName = "Daniel", LastName = "Day-Lewis", YearBorn = 1957 },
                                                              new Person { FirstName = "Sally", LastName = "Field", YearBorn = 1946 },
                                                              new Person { FirstName = "David", LastName = "Strathairn", YearBorn = 1949 },
                                                              new Person { FirstName = "Joseph", LastName = "Gordon-Levitt", YearBorn = 1981 },
                                                              new Person { FirstName = "James", LastName = "Spader", YearBorn = 1960 }
                                                          };

        private readonly List<Person> _personList;

        [Fact]
        public void Columns_should_be_autogenerated_if_no_columns_are_specified()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var package = _personList.ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Columns.Should().Be(2);
        }

        [Fact]
        public void Header_texts_should_be_same_as_defined_on_ExcelTableColumn_attribute()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var package = _personList.ToExcelPackage(true);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Cells[1, 1, 1, 1].Value.Should().Be("LastName");
            package.GetWorksheet(worksheetIndex).Cells[1, 2, 1, 2].Value.Should().Be("Year of Birth");
            package.GetWorksheet(worksheetIndex).Dimension.Rows.Should().Be(_personList.Count + 1);
        }

        [Fact]
        public void Multiple_lists_should_create_multiple_worksheets()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<Person> pre50 = _personList.Where(x => x.YearBorn < 1950).ToList();
            List<Person> post50 = _personList.Where(x => x.YearBorn >= 1950).ToList();

            var worksheetPre50Index = 0;
            var worksheetPost50Index = 1;
#if NETFRAMEWORK
            worksheetPre50Index = 1;
            worksheetPost50Index = 2;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = pre50
                .ToWorksheet("< 1950")
                .WithColumn(x => x.FirstName, "First Name")
                .WithColumn(x => x.LastName, "Last Name")
                .WithColumn(x => x.YearBorn, "Year of Birth")
                .WithTitle("< 1950")
                .NextWorksheet(post50, "> 1950")
                .WithColumn(x => x.LastName, "Last Name")
                .WithColumn(x => x.YearBorn, "Year of Birth")
                .WithTitle("> 1950")
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Workbook.Worksheets.Count.Should().Be(2);
            package.GetWorksheet(worksheetPre50Index).Dimension.Rows.Should().Be(pre50.Count + 2);
            package.GetWorksheet(worksheetPre50Index).Dimension.Columns.Should().Be(3);
            package.GetWorksheet(worksheetPost50Index).Dimension.Rows.Should().Be(post50.Count + 2);
            package.GetWorksheet(worksheetPost50Index).Dimension.Columns.Should().Be(2);
        }

        [Fact]
        public void Multiple_lists_should_create_multiple_worksheets_without_header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<Person> pre50 = _personList.Where(x => x.YearBorn < 1950).ToList();
            List<Person> post50 = _personList.Where(x => x.YearBorn >= 1950).ToList();

            var worksheetPre50Index = 0;
            var worksheetPost50Index = 1;
#if NETFRAMEWORK
            worksheetPre50Index = 1;
            worksheetPost50Index = 2;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = pre50
                .ToWorksheet("< 1950")
                .WithConfiguration(configuration =>
                                   configuration.WithColumnConfiguration(x => x.SetFontColor(Color.Purple))
                                                .WithHeaderConfiguration(x => x.SetFont(new Font("Arial", 13, FontStyle.Bold))
                                                                               .SetFontColor(Color.White)
                                                                               .SetBackgroundColor(Color.Black))
                                                .WithHeaderRowConfiguration(x => x.BorderAround(ExcelBorderStyle.Thin)
                                                                                  .SetFontName("Verdana"))
                                                .WithCellConfiguration((x, y) =>
                                                                       {
                                                                           x.SetFont(new Font("Times New Roman", 13));
                                                                           y.YearBorn = y.YearBorn % 2 == 0 ? y.YearBorn : 1990;
                                                                       })
                                                .WithTitleConfiguration(x => x.SetBackgroundColor(Color.Yellow))
                )
                .WithTitle("< 1950")
                .NextWorksheet(post50, "> 1950")
                .WithConfiguration(configuration =>
                                   {
                                       configuration.WithColumnConfiguration(x => x.SetFontColor(Color.Black))
                                                    .WithHeaderConfiguration(x =>
                                                                             {
                                                                                 x.Style.Font.Bold = true;
                                                                                 x.Style.Font.Size = 11;
                                                                                 x.SetFontColor(Color.White);
                                                                                 x.SetBackgroundColor(Color.Black, ExcelFillStyle.Solid);
                                                                             })
                                                    .WithHeaderRowConfiguration(x => x.BorderAround(ExcelBorderStyle.Thin)
                                                                                      .SetFontName("Verdana"))
                                                    .WithCellConfiguration((x, y) =>
                                                                           {
                                                                               x.SetFontName("Times New Roman");
                                                                               y.YearBorn = y.YearBorn % 2 != 0 ? y.YearBorn : 1990;
                                                                           });
                                   })
                .WithoutHeader()
                .WithTitle("> 1950")
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Workbook.Worksheets.Count.Should().Be(2);
            package.GetWorksheet(worksheetPre50Index).Dimension.Rows.Should().Be(pre50.Count + 2);
            package.GetWorksheet(worksheetPre50Index).Dimension.Columns.Should().Be(2);
            package.GetWorksheet(worksheetPost50Index).Dimension.Rows.Should().Be(post50.Count + 1);
            package.GetWorksheet(worksheetPost50Index).Dimension.Columns.Should().Be(2);
            package.GetWorksheet(worksheetPre50Index).Cells[2, 1, 1, 1].Style.Font.Size.Should().Be(13);
            package.GetWorksheet(worksheetPre50Index).Cells[2, 1, 1, 1].Style.Font.Name.Should().Be("Verdana");
            package.GetWorksheet(worksheetPre50Index).Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should().Be(string.Format("{0:X8}", Color.Yellow.ToArgb() & 0xFFFFFFFF));
        }

        [Fact]
        public void Rowcount_should_match_listcount_plus_header_if_not_title()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList.ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Rows.Should().Be(_personList.Count + 1);
        }

        [Fact]
        public void Rowcount_should_match_listcount_plus_header_plus_one_title()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList
                .ToWorksheet("Actors")
                .WithTitle("Actors")
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Rows.Should().Be(_personList.Count + 2);
        }

        [Fact]
        public void Rowcount_should_match_listcount_plus_header_plus_two_titles()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList
                .ToWorksheet("Actors")
                .WithTitle("Actors")
                .WithTitle("In the movie Lincoln")
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Rows.Should().Be(_personList.Count + 3);
        }

        [Fact]
        public void Rowcount_should_match_listcount_with_header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList.ToExcelPackage(true);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Rows.Should().Be(_personList.Count + 1);
        }

        [Fact]
        public void Rowcount_should_match_listcount_without_header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList.ToExcelPackage(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Rows.Should().Be(_personList.Count);
        }

        [Fact]
        public void Should_convert_a_byte_array_into_an_excelPackage()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer1 = ExcelPackage1.GetAsByteArray();
            var buffer2 = new byte[] { };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package1 = buffer1.AsExcelPackage();
            Action act = () => buffer2.AsExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package1.Should().NotBeNull();
            package1.Workbook.Worksheets.Count.Should().Be(ExcelPackage1.Workbook.Worksheets.Count);
            package1.GetAllTables().Count().Should().Be(ExcelPackage1.GetAllTables().Count());

            act.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void Should_convert_a_byte_array_into_an_excelPackage_with_password()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer1 = ExcelPackage1.GetAsByteArray("Test1234");
            var buffer2 = new byte[] { };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = buffer1.AsExcelPackage("Test1234");
            Action act = () => buffer2.AsExcelPackage("test");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Should().NotBeNull();
            package.Workbook.Worksheets.Count.Should().Be(ExcelPackage1.Workbook.Worksheets.Count);
            package.GetAllTables().Count().Should().Be(ExcelPackage1.GetAllTables().Count());

            act.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void Should_convert_a_stream_into_an_excelPackage()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Stream stream = new MemoryStream(ExcelPackage1.GetAsByteArray());

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = stream.AsExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Should().NotBeNull();
            package.Workbook.Worksheets.Count.Should().Be(ExcelPackage1.Workbook.Worksheets.Count);
            package.GetAllTables().Count().Should().Be(ExcelPackage1.GetAllTables().Count());
        }

        [Fact]
        public void Should_convert_a_stream_into_an_excelPackage_with_correct_password()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer = ExcelPackage1.GetAsByteArray("Test1234");
            var stream = new MemoryStream(buffer);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = stream.AsExcelPackage("Test1234");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Should().NotBeNull();
            package.Workbook.Worksheets.Count.Should().Be(ExcelPackage1.Workbook.Worksheets.Count);
            package.GetAllTables().Count().Should().Be(ExcelPackage1.GetAllTables().Count());
        }

        [Fact]
        public void Should_convert_a_worksheet_to_byte_array_of_xlsx_file()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<Person> post50 = _personList.Where(x => x.YearBorn > 1950).ToList();
            string tempFile = Path.Combine(Path.GetTempPath(), "persons_post50.xlsx");
            WorksheetWrapper<Person> worksheetWrapper = post50.ToWorksheet("> 1950");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            byte[] result = worksheetWrapper.ToXlsx();
            File.WriteAllBytes(tempFile, result);
            var excelPackage = new ExcelPackage(new FileInfo(tempFile));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            excelPackage.Workbook.Worksheets.Any().Should().Be(true);
            excelPackage.Workbook.Worksheets.First().Dimension.Rows.Should().Be(4);
        }

        [Fact]
        public void Should_convert_list_of_objects_to_byte_array_of_xlsx_file_with_header_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<Person> pre50 = _personList.Where(x => x.YearBorn < 1950).ToList();
            string tempFile = Path.Combine(Path.GetTempPath(), "persons_pre50.xlsx");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            byte[] result = pre50.ToXlsx(true);
            File.WriteAllBytes(tempFile, result);
            var excelPackage = new ExcelPackage(new FileInfo(tempFile));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            excelPackage.Workbook.Worksheets.Any().Should().Be(true);
            excelPackage.Workbook.Worksheets.First().Should().NotBe(null);
            excelPackage.Workbook.Worksheets.First().Dimension.Rows.Should().Be(3);
            excelPackage.Workbook.Worksheets.First().Cells[1, 1, 1, 1].Text.Should().Be("LastName");
            excelPackage.Workbook.Worksheets.First().Cells[1, 2, 1, 2].Text.Should().Be("Year of Birth");
        }

        [Fact]
        public void Should_convert_list_of_objects_to_byte_array_of_xlsx_file_without_header_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<Person> pre50 = _personList.Where(x => x.YearBorn < 1950).ToList();
            string tempFile = Path.Combine(Path.GetTempPath(), "persons_pre50.xlsx");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            byte[] result = pre50.ToXlsx(false);
            File.WriteAllBytes(tempFile, result);
            var excelPackage = new ExcelPackage(new FileInfo(tempFile));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            excelPackage.Workbook.Worksheets.Any().Should().Be(true);
            excelPackage.Workbook.Worksheets.First().Should().NotBe(null);
            excelPackage.Workbook.Worksheets.First().Dimension.Rows.Should().Be(2);
            excelPackage.Workbook.Worksheets.First().Cells[1, 1, 1, 1].Text.Should().Be("Field");
            excelPackage.Workbook.Worksheets.First().Cells[1, 2, 1, 2].Text.Should().Be("1946");
        }

        [Fact]
        public void Should_convert_multiple_worksheets_to_byte_array_of_xlsx_file()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<Person> pre50 = _personList.Where(x => x.YearBorn < 1950).ToList();
            List<Person> post50 = _personList.Where(x => x.YearBorn >= 1950).ToList();

            var worksheetPre50Index = 0;
            var worksheetPost50Index = 1;
#if NETFRAMEWORK
            worksheetPre50Index = 1;
            worksheetPost50Index = 2;
#endif

            WorksheetWrapper<Person> worksheetWrapper = pre50
                .ToWorksheet("< 1950")
                .WithTitle("< 1950")
                .WithConfiguration(x => x.WithTitleConfiguration(c => c.SetBackgroundColor(Color.AliceBlue)))
                .NextWorksheet(post50, "> 1950")
                .WithTitle("> 1950", c => c.SetBackgroundColor(Color.Aquamarine));

            string tempFile = Path.Combine(Path.GetTempPath(), "persons_pre50_post50.xlsx");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            byte[] result = worksheetWrapper.ToXlsx();
            File.WriteAllBytes(tempFile, result);
            var excelPackage = new ExcelPackage(new FileInfo(tempFile));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            excelPackage.Workbook.Worksheets.Count().Should().Be(2);
            excelPackage.GetWorksheet(worksheetPre50Index).Dimension.Rows.Should().Be(4);
            excelPackage.GetWorksheet(worksheetPost50Index).Dimension.Rows.Should().Be(5);
            excelPackage.GetWorksheet(worksheetPre50Index).Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.AliceBlue.ToArgb() & 0xFFFFFFFF:X8}");
            excelPackage.GetWorksheet(worksheetPost50Index).Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Aquamarine.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_not_convert_a_byte_array_into_an_excelPackage_with_empty_password()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer = ExcelPackage1.GetAsByteArray();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => buffer.AsExcelPackage("");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void Should_not_convert_a_byte_array_into_an_excelPackage_with_wrong_password()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer = ExcelPackage1.GetAsByteArray("Test1234");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => buffer.AsExcelPackage("test12345");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<SecurityException>();
        }

        [Fact]
        public void Should_not_convert_a_stream_into_an_excelPackage_with_incorrect_password()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer = ExcelPackage1.GetAsByteArray("Test1234");
            var stream = new MemoryStream(buffer);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => stream.AsExcelPackage("test1234");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<SecurityException>();
        }

        [Fact]
        public void Worksheet_should_match_specified_int_columns()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList
                .ToWorksheet("Actors")
                .WithColumn(x => x.YearBorn, "Year Born", c => c.SetBackgroundColor(Color.Azure))
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Columns.Should().Be(1);

            for (var i = 0; i < _personList.Count; i++)
            {
                package.GetWorksheet(worksheetIndex).Cells[i + 2, 1].Value.Should().Be(_personList[i].YearBorn);
                package.GetWorksheet(worksheetIndex).Cells[i + 2, 1].Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Azure.ToArgb() & 0xFFFFFFFF:X8}");
            }
        }

        [Fact]
        public void Worksheet_should_match_specified_string_columns()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            var worksheetIndex = 0;
#if NETFRAMEWORK
            worksheetIndex = 1;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList
                .ToWorksheet("Actors")
                .WithColumn(x => x.LastName, "Last Name")
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.GetWorksheet(worksheetIndex).Dimension.Columns.Should().Be(1);
            for (var i = 0; i < _personList.Count; i++)
            {
                package.GetWorksheet(worksheetIndex).Cells[i + 2, 1].Text.Should().Be(_personList[i].LastName);
            }
        }

        [Fact]
        public void Should_add_multiple_titles()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package;
            const string worksheetName = "Actors";

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            package = _personList
                .ToWorksheet(worksheetName)
                .WithTitle("title 1")
                .WithTitle("title 2")
                .WithTitle("title 3")
                .ToExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            var worksheet = package.GetWorksheet(worksheetName);
            worksheet.Cells[1, 1].Value.Should().Be("title 1");
            worksheet.Cells[2, 1].Value.Should().Be("title 2");
            worksheet.Cells[3, 1].Value.Should().Be("title 3");
        }
    }
}