﻿#nullable disable
using System.Data;
using System.Drawing;

using OLT.Extensions.EPPlus.Exceptions;
using OLT.Extensions.EPPlus.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace OLT.Extensions.EPPlus.Tests
{
    public class ExcelWorksheetExtensionsTests : TestBase
    {
        [Fact]
        public void Should_add_an_header_without_configuration()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddHeader("NewBarcode", "NewQuantity", "NewUpdatedDate");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Dimension.End.Row.Should().Be(5);
            worksheet.Cells[1, 1, 1, 1].Value.Should().Be("NewBarcode");
            worksheet.Cells[1, 2, 1, 2].Value.Should().Be("NewQuantity");
            worksheet.Cells[1, 3, 1, 3].Value.Should().Be("NewUpdatedDate");
            worksheet.Cells[2, 1, 2, 1].Value.Should().Be("Barcode");
        }

        [Fact]
        public void Should_add_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");
            Color color = Color.AntiqueWhite;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddHeader(x => x.SetBackgroundColor(color), "NewBarcode", "NewQuantity", "NewUpdatedDate");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Dimension.End.Row.Should().Be(5);
            worksheet.Cells[1, 1, 1, 1].Value.Should().Be("NewBarcode");
            worksheet.Cells[1, 2, 1, 2].Value.Should().Be("NewQuantity");
            worksheet.Cells[1, 3, 1, 3].Value.Should().Be("NewUpdatedDate");

            worksheet.Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should().Be($"{color.ToArgb() & 0xFFFFFFFF:X8}");
            worksheet.Cells[2, 1, 2, 1].Value.Should().Be("Barcode");
        }

        [Fact]
        public void Should_add_line_with_configuration()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddLine(5, configureCells => configureCells.SetBackgroundColor(Color.Yellow), "barcode123", 5,
                              DateTime.UtcNow);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>(configuration => configuration
                                                                                                     .WithoutHeaderRow()
                                                                                                     .SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells[5, 1].Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Yellow.ToArgb() & 0xFFFFFFFF:X8}");
            list.Count().Should().Be(5);
        }

        [Fact]
        public void Should_add_line_without_configuration()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddLine(5, "barcode123", 5, DateTime.UtcNow);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_add_objects_with_parameters()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");
            DateTime dateTime = DateTime.MaxValue;

            var stocks = new List<StocksNullable>
                         {
                             new StocksNullable
                             {
                                 Barcode = "barcode123",
                                 Quantity = 5,
                                 UpdatedDate = dateTime
                             }
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddObjects(stocks, 5, _ => _.Barcode, _ => _.Quantity, _ => _.UpdatedDate);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
            list.Last().Barcode.Should().Be("barcode123");
            list.Last().Quantity.Should().Be(5);
            list.Last().UpdatedDate.HasValue.Should().BeTrue();
            list.Last().UpdatedDate?.Date.Should().Be(dateTime.Date);
            list.Last().UpdatedDate?.Hour.Should().Be(dateTime.Hour);
        }

        [Fact]
        public void Should_add_objects_with_start_row_and_column_index_without_parameters()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");

            var stocks = new List<StocksNullable>
                         {
                             new StocksNullable
                             {
                                 Barcode = "barcode123",
                                 Quantity = 5,
                                 UpdatedDate = DateTime.MaxValue
                             }
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddObjects(stocks, 5, 3);
            IEnumerable<StocksNullable> list =
                worksheet.ToList<StocksNullable>(configuration => configuration.SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_cannot_add_objects_with_null_property_selectors()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");
            var stocks = new List<StocksNullable>
                         {
                             new StocksNullable
                             {
                                 Barcode = "barcode123",
                                 Quantity = 5,
                                 UpdatedDate = DateTime.MaxValue
                             }
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => worksheet.AddObjects(stocks, 5, null);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void Should_check_and_throw_exception_if_column_value_is_wrong_on_specified_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () => worksheet.CheckColumnAndThrow(2, 3, "Barcode", "Barcode column is missing");

            Action action2 = () => worksheet.CheckColumnAndThrow(2, 1, "Barcode");

            Action action3 = () => worksheet.CheckColumnAndThrow(3, 14, "Barcode");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<ExcelValidationException>().And.Message.Should().Be("Barcode column is missing");
            action2.Should().NotThrow();
            action3.Should().Throw<ExcelValidationException>();
        }

        [Fact]
        public void Should_check_and_throw_if_duplicated_column_found_on_a_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage2.GetWorksheet("RandomOrderedColumns");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action act1 = () => worksheet1.CheckAndThrowIfDuplicatedColumnsFound(1);
            Action act2 = () => worksheet1.CheckAndThrowIfDuplicatedColumnsFound(1, "'{0}' column is duplicated (rowIndex: {1})");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            act1.Should().Throw<ExcelValidationException>().WithMessage("'Barcode' column is duplicated on 1. row.");
            act2.Should().Throw<ExcelValidationException>().WithMessage("'Barcode' column is duplicated (rowIndex: 1)");
        }

        [Fact]
        public void Should_check_and_throw_if_duplicated_column_found_on_a_row_with_object()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage2.GetWorksheet("RandomOrderedColumns");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action act1 = () => worksheet1.CheckAndThrowIfDuplicatedColumnsFound<UnorderedBarcodeAndQuantity>(1);
            Action act2 = () => worksheet1.CheckAndThrowIfDuplicatedColumnsFound<UnorderedBarcodeAndQuantity>(1, "'{0}' column is duplicated (rowIndex: {1})");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            act1.Should().Throw<ExcelValidationException>().WithMessage("'Barcode' column is duplicated on 1. row.");
            act2.Should().Throw<ExcelValidationException>().WithMessage("'Barcode' column is duplicated (rowIndex: 1)");
        }

        [Fact]
        public void Should_check_if_a_cell_is_null_or_empty_on_given_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result1 = worksheet.IsCellEmpty(3, 4);
            bool result2 = worksheet.IsCellEmpty(2, 1);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result1.Should().BeTrue();
            result2.Should().BeFalse();
        }

        [Fact]
        public void Should_check_if_duplicated_column_exists_on_a_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage2.GetWorksheet("RandomOrderedColumns");
            ExcelWorksheet worksheet2 = ExcelPackage2.GetWorksheet("EmptyWorksheet");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool duplicatedColumn = worksheet1.IsColumnDuplicatedOnRow(1, "Barcode");
            bool notDuplicatedColumn = worksheet1.IsColumnDuplicatedOnRow(1, "Quantity");
            bool notfoundColumn = worksheet1.IsColumnDuplicatedOnRow(1, "Barcode1asdacfddsd");
            bool emptyRow = worksheet2.IsColumnDuplicatedOnRow(2, "empty");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            duplicatedColumn.Should().BeTrue();
            notDuplicatedColumn.Should().BeFalse();
            notfoundColumn.Should().BeFalse();
            emptyRow.Should().BeFalse();
        }

        [Fact]
        public void Should_convert_to_datatable_with_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = ExcelPackage1.GetWorksheet("TEST5").ToDataTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(3, "We have 3 records");
        }

        [Fact]
        public void Should_convert_to_datatable_without_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var dataTable = ExcelPackage1.GetWorksheet("TEST5").ToDataTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(4, "We have 4 records");
        }

        [Fact]
        public void Should_convert_worksheet_to_list()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage1.GetWorksheet("TEST4");
            ExcelWorksheet worksheet2 = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> list1 = worksheet1.ToList<StocksNullable>(configuration =>
                                                                           {
                                                                               configuration.SkipCastingErrors();
                                                                               configuration.WithoutHeaderRow();
                                                                           });

            List<StocksNullable> list2 = worksheet2.ToList<StocksNullable>(configuration =>
                                                                           {
                                                                               configuration.SkipCastingErrors();
                                                                               configuration.WithoutHeaderRow();
                                                                           });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count.Should().Be(4, "Should have four");
            list2.Count.Should().Be(4, "Should have four");
        }

        [Fact]
        public void Should_delete_a_column_by_using_header_text()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");
            const string columnName = "Quantity";

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.DeleteColumn(columnName);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.GetColumns(1).Any(x => x.Value == columnName).Should().BeFalse();
            worksheet.GetValuedDimension().End.Column.Should().Be(2);
            worksheet.Cells[2, 2, 2, 2].Text.Should().Be("UpdatedDate");
        }

        [Fact]
        public void Should_delete_columns_by_given_header_text()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");
            const string columnName = "Quantity";

            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            worksheet.ChangeCellValue(2, valuedDimension.End.Column + 1, columnName);
            worksheet.ChangeCellValue(2, valuedDimension.End.Column + 2, columnName);
            worksheet.ChangeCellValue(2, valuedDimension.End.Column + 3, columnName);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.DeleteColumns(columnName);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.GetColumns(2).Any(x => x.Value == columnName).Should().BeFalse();
            worksheet.GetValuedDimension().End.Column.Should().Be(2);
            worksheet.Cells[2, 2, 2, 2].Text.Should().Be("UpdatedDate");
        }

        [Fact]
        public void Should_find_formulas_on_a_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage1.GetWorksheet("TEST5");
            worksheet1.Cells[18, 2, 18, 2].Formula = "=SUM(B2:B4)";

            ExcelWorksheet worksheet2 = ExcelPackage1.GetWorksheet("TEST4");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result1 = worksheet1.HasAnyFormula();
            bool result2 = worksheet2.HasAnyFormula();
            Action action1 = () => { worksheet1.CheckAndThrowIfThereIsAnyFormula("First worksheet has formulas."); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result1.Should().BeTrue();
            result2.Should().BeFalse();
            action1.Should().Throw<ExcelValidationException>().WithMessage("First worksheet has formulas.");
        }

        [Fact]
        public void Should_get_as_Excel_table_with_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            List<StocksNullable> listOfStocks = excelTable.ToList<StocksNullable>();
            listOfStocks.Count.Should().Be(3);
        }

        [Fact]
        public void Should_get_as_Excel_table_without_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            List<StocksNullable> listOfStocks =
                excelTable.ToList<StocksNullable>(configuration => configuration.SkipCastingErrors());
            listOfStocks.Count.Should().Be(4);
        }

        [Fact]
        public void Should_get_columns_of_given_row_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<KeyValuePair<int, string>> results = worksheet.GetColumns(2).ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count().Should().Be(3);
            results.First().Value.Should().Be("Barcode");
        }

        [Fact]
        public void Should_get_data_bounds_of_worksheet_with_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBounds = excelWorksheet.GetDataBounds();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBounds.Rows.Should().Be(3);
            dataBounds.Columns.Should().Be(3);
        }

        [Fact]
        public void Should_get_data_bounds_of_worksheet_without_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBounds = excelWorksheet.GetDataBounds(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBounds.Rows.Should().Be(4);
            dataBounds.Columns.Should().Be(3);
        }

        [Fact]
        public void Should_get_empty_list_if_table_is_empty()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.Workbook.Worksheets["TEST7"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> results = worksheet.ToList<StocksNullable>(configuration =>
                                                                            configuration.Intercept((item, row) => { item.Barcode = item.Barcode?.Trim(); })
                );

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().Be(0);
        }

        [Fact]
        public void Should_get_worksheet_as_enumerable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage1.GetWorksheet("TEST4");
            ExcelWorksheet worksheet2 = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> list1 = worksheet1.AsEnumerable<StocksNullable>(configuration =>
                                                                                        {
                                                                                            configuration.SkipCastingErrors();
                                                                                            configuration.SkipValidationErrors();
                                                                                            configuration.WithoutHeaderRow();
                                                                                        });

            IEnumerable<StocksNullable> list2 = worksheet2.AsEnumerable<StocksNullable>(configuration =>
                                                                                        {
                                                                                            configuration.SkipCastingErrors();
                                                                                            configuration.SkipValidationErrors();
                                                                                            configuration.WithoutHeaderRow();
                                                                                        });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }

        [Fact]
        public void Should_parse_datetime_value_as_correctly_even_if_field_has_a_custom_format()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> nullableStocks = worksheet.ToList<StocksNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            nullableStocks[0].UpdatedDate.HasValue.Should().Be(true);
            nullableStocks[0].UpdatedDate?.Date.Should().Be(new DateTime(2017, 08, 08));

            nullableStocks[1].UpdatedDate.HasValue.Should().Be(true);
            nullableStocks[1].UpdatedDate?.Should().Be(new DateTime(2016, 11, 03, 01, 30, 53));
        }

        [Fact]
        public void Should_read_worksheet_without_considering_orders_of_columns()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage2.GetWorksheet("RandomOrderedColumns");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<UnorderedBarcodeAndQuantity> result = worksheet.ToList<UnorderedBarcodeAndQuantity>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result.Count.Should().Be(3);
            result[0].Barcode.Should().Be("barcode1");
            result[0].Quantity.Should().Be(11);
            result[0].UpdatedDate.Date.Should().Be(new DateTime(2018, 06, 25));

            result[1].Barcode.Should().Be("barcode2");
            result[1].Quantity.Should().Be(22);
            result[1].UpdatedDate.Date.Should().Be(new DateTime(2018, 06, 26));

            result[2].Barcode.Should().Be("barcode3");
            result[2].Quantity.Should().Be(33);
            result[2].UpdatedDate.Date.Should().Be(new DateTime(2018, 06, 27));
        }

        [Fact]
        public void Should_throw_an_exception_when_columns_of_worksheet_not_matched_with_ExcelTableAttribute()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage1.GetWorksheet("TEST6");
            ExcelWorksheet emptySheet1 = ExcelPackage1.GetWorksheet("EmptySheet");
            ExcelWorksheet emptySheet2 = ExcelPackage1.GetWorksheet("EmptySheet");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () => worksheet1.CheckHeadersAndThrow<NamedMap>(2, "The {0}.column of worksheet should be '{1}'.");
            Action action2 = () => worksheet1.CheckHeadersAndThrow<NamedMap>(2);
            Action action3 = () => worksheet1.CheckHeadersAndThrow<StocksNullable>(2);
            Action action4 = () => worksheet1.CheckHeadersAndThrow<Car>(2);

            Action actionForEmptySheet1 = () => emptySheet1.CheckHeadersAndThrow<StocksValidation>(1, "The {0}.column of worksheet should be '{1}'.");
            Action actionForEmptySheet2 = () => emptySheet2.CheckHeadersAndThrow<Cars>(1, "The {0}.column of worksheet should be '{1}'.");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<ExcelValidationException>().And.Message.Should().Be("The 1.column of worksheet should be 'Name'.");
            action2.Should().Throw<ExcelValidationException>().And.Message.Should().Be("The 1. column of worksheet should be 'Name'.");
            action3.Should().NotThrow<ExcelValidationException>();
            action4.Should().Throw<InvalidOperationException>();

            actionForEmptySheet1.Should().Throw<ExcelValidationException>().And.Message.Should().Be("The 1.column of worksheet should be 'Barcode'.");
            actionForEmptySheet2.Should().Throw<ExcelValidationException>().And.Message.Should().Be("The 1.column of worksheet should be 'LicensePlate'.");
        }

        [Fact]
        public void Should_throw_Excel_validation_exception_if_worksheet_does_not_have_valued_dimension()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("EmptySheet");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<Cars> cars = worksheet.ToList<Cars>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            cars.Count.Should().Be(0);
        }

        [Fact]
        public void Should_throw_exception_when_occured_validation_error()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { worksheet.ToList<StocksValidation>(); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ExcelValidationException>()
                  .WithMessage("Please enter a value bigger than 10")
                  .And.Args.ColumnName.Should().Be("Quantity");
        }

        [Fact]
        public void Should_valued_dimension_be_A2C5()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            valuedDimension.Address.Should().Be("A2:C5");
            valuedDimension.Start.Column.Should().Be(1);
            valuedDimension.Start.Row.Should().Be(2);
            valuedDimension.End.Column.Should().Be(3);
            valuedDimension.End.Row.Should().Be(5);
        }

        [Fact]
        public void Should_valued_dimension_be_correct()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            const string expectedAddress = "E9:G13";
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet("TEST4");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            valuedDimension.Address.Should().Be(expectedAddress);
            valuedDimension.Start.Column.Should().Be(5);
            valuedDimension.Start.Row.Should().Be(9);
            valuedDimension.End.Column.Should().Be(7);
            valuedDimension.End.Row.Should().Be(13);
        }


        [Fact]
        public void Should_check_columns_of_given_row_whether_it_contains_same_values_with_excel_exportable_object_properties()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage2.GetWorksheet("RandomOrderedColumns");
            ExcelWorksheet emptySheet1 = ExcelPackage1.GetWorksheet("EmptySheet");
            ExcelWorksheet emptySheet2 = ExcelPackage1.GetWorksheet("EmptySheet");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () => worksheet1.CheckExistenceOfColumnsAndThrow<NamedMap>(2, "'{0}' column is not found on the Excel.");
            Action action2 = () => worksheet1.CheckExistenceOfColumnsAndThrow<NamedMap>(2);
            Action action3 = () => worksheet1.CheckExistenceOfColumnsAndThrow<UnorderedBarcodeAndQuantity>(1);
            Action action4 = () => worksheet1.CheckExistenceOfColumnsAndThrow<Car>(2);

            Action actionForEmptySheet1 = () => emptySheet1.CheckExistenceOfColumnsAndThrow<StocksValidation>(1, "'{0}' column is not found on the worksheet.");
            Action actionForEmptySheet2 = () => emptySheet2.CheckExistenceOfColumnsAndThrow<Cars>(1, "'{0}' column is not found on the worksheet.");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<ExcelValidationException>().And.Message.Should().Be("'Name' column is not found on the Excel.");
            action2.Should().Throw<ExcelValidationException>().And.Message.Should().Be("'Name' column is not found on the worksheet.");
            action3.Should().NotThrow<ExcelValidationException>();
            action4.Should().Throw<InvalidOperationException>();

            actionForEmptySheet1.Should().Throw<ExcelValidationException>().And.Message.Should().Be("'Barcode' column is not found on the worksheet.");
            actionForEmptySheet2.Should().Throw<ExcelValidationException>().And.Message.Should().Be("'LicensePlate' column is not found on the worksheet.");
        }

        [Fact]
        public void Should_throw_exception_if_columnname_and_columnindex_not_defined_and_column_missing_on_Excel()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = ExcelPackage2.GetWorksheet("RandomOrderedColumns");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action act = () => worksheet1.ToList<DefaultMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            act.Should().Throw<ExcelValidationException>().And.Message.Should().Be("'Name' column could not found on the worksheet.");
        }


        [Fact]
        public void Should_not_throw_exception_if_a_column_is_marked_as_Optional_and_missing()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheetWithOptionalColumns = ExcelPackage1.GetWorksheet("WithOptionalFields");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var results = worksheetWithOptionalColumns.ToList<ExcelWithOptionalFields>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().Be(2);
            results.All(x => x.MissingColumn1 == default).Should().BeTrue();
            results.All(x => x.MissingColumn2 == null).Should().BeTrue();
            results.All(x => x.MissingColumn3 == null).Should().BeTrue();
        }
    }
}
