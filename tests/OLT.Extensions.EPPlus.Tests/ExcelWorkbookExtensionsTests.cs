using OLT.Extensions.EPPlus.Style;
using AwesomeAssertions;
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

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action act1 = () => ExcelWorkbookStyleExtensions.CreateNamedStyle(workbook1, "Style1", act => act.SetTextVertical());

            
            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            act1.Should().Throw<InvalidOperationException>().WithMessage("The Excel package already has a style with the name of 'Style1'");
        }


        [Fact]
        public void Should_check_named_style_found()
        {

            ExcelWorkbook workbook1 = ExcelPackage2.Workbook;
            ExcelWorkbookStyleExtensions.CreateNamedStyle(workbook1, "Style1", act => act.SetFontAsBold());


            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action act1 = () => ExcelWorkbookStyleExtensions.CreateNamedStyleIfNotExists(workbook1, "Style1", act => act.SetTextVertical());


            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            act1.Should().NotThrow<InvalidOperationException>();
        }
    }
}
