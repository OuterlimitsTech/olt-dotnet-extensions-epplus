using OfficeOpenXml;
using OfficeOpenXml.Style;
using OLT.Core;
using System;
using System.Collections.Generic;
using System.IO;

namespace OLT.EPPlus.Tests.Assets
{
    public class TestFileBuilder 
    {
        public IOltFileBase64 Build(TestFileBuilderRequest request)
        {
            using var excelPackage = new ExcelPackage();
            var worksheet = excelPackage.Workbook.Worksheets.Add($"Count {request.Data.Count:N0}");

            worksheet.Write(ExcelColumns, 1);
            var rows = new List<IOltExcelRowWriter>();

            request.Data.ForEach(service =>
            {
                var cells = new List<IOltExcelCellWriter>
                {
                    new OltExcelCellWriter(service.Email),
                    new OltExcelCellWriter(service.First),
                    new OltExcelCellWriter(service.Last),
                };


                rows.Add(new OltExcelRowWriter { Cells = cells });
                rows.Add(new OltExcelRowEmpty());
            });


            var rowIdx = 2;
            rows.ForEach(row =>
            {
                rowIdx = row.Write(worksheet, rowIdx);
            });

            return new OltFileBase64
            {
                FileName = $"{nameof(TestFileBuilder)} [{DateTimeOffset.Now:s}].xlsx",
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileBase64 = Convert.ToBase64String(excelPackage.GetAsByteArray())
            };
        }


        #region [ Excel ]

        private static List<IOltExcelColumn> ExcelColumns
        {
            get
            {
                var columns = new List<IOltExcelColumn>
                {
                    new OltExcelColumn
                    {
                        Heading = "Person Id",
                        Width = 22.5m,
                        Style = new OltExcelCellStyle
                        {
                            HorizontalAlignment = ExcelHorizontalAlignment.Center
                        }
                    },
                    new OltExcelColumn
                    {
                        Heading = "Email",
                        Width = 15.5m,
                        Style = new OltExcelCellStyle
                        {
                            WrapText = true,
                            HorizontalAlignment = ExcelHorizontalAlignment.Left
                        }
                    },
                    new OltExcelColumn
                    {
                        Heading = "First Name",
                        Width = 15.5m,
                        Style = new OltExcelCellStyle
                        {
                            WrapText = true,
                            HorizontalAlignment = ExcelHorizontalAlignment.Justify
                        }
                    },
                    new OltExcelColumn
                    {
                        Heading = "Last Name",
                        Width = 15.5m,
                        Style = new OltExcelCellStyle
                        {
                            WrapText = true,
                            HorizontalAlignment = ExcelHorizontalAlignment.Right
                        }
                    },

                };
                return columns;
            }
        }

        #endregion


        public static string BuildTempPath()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), $"OLT_UnitTest_EPPlus_{Guid.NewGuid()}");
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }
            return tempDir;
        }

        public static string BuildTempPath(string rootDir)
        {
            var tempDir = Path.Combine(Path.GetTempPath(), rootDir, $"OLT_UnitTest_EPPlus_{Guid.NewGuid()}");
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }
            return tempDir;
        }
    }
}