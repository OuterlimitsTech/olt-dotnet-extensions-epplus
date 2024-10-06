using FluentAssertions;

namespace OLT.Extensions.EPPlus.Tests
{
    public class ExcelHelperTests : TestBase
    {
        [Theory]
        [InlineData("A", 1)]
        [InlineData("AG", 33)]
        [InlineData("ZZ", 702)]
        [InlineData("XZ", 650)]
        public void ColumnLetterToColumnIndex(string columnLetter, int expected)
        {
            var result = ExcelHelper.ColumnLetterToColumnIndex(columnLetter);
            Assert.Equal(result, expected);
        }

        [Fact]
        public void ColumnLetterToColumnIndexNull()
        {
            Action action = () => ExcelHelper.ColumnLetterToColumnIndex(null);
            action.Should()
                .Throw<ArgumentNullException>();
        }

        [Theory]
        [InlineData("AZZ")]
        [InlineData("A1")]
        [InlineData("")]
        public void ColumnLetterToColumnIndexOutofRange(string columnLetter)
        {
            Action action = () => ExcelHelper.ColumnLetterToColumnIndex(columnLetter);
            action.Should()
                .Throw<ArgumentOutOfRangeException>()
                .WithMessage("Must only be 1-2 alpha characters (Parameter 'columnLetter')");
        }
    }
}
