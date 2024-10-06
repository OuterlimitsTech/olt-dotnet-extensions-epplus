namespace OLT.Extensions.EPPlus.Tests
{
    public class ExcelCsvConverterTests : TestBase
    {

        [Fact]
        public void ConvertToCsvTest()
        {
            string exportDirectory = String.Empty;
            var files = new List<string>();

            try
            {
                exportDirectory = BuildTempPath();

                using var package = ImportTestExcelPackage;
                var worksheets = OLT.Extensions.EPPlus.ExcelCsvConverter.ToCsv(package);
             
                worksheets.ForEach(item =>
                {
                    var fileName = Path.Combine(exportDirectory, $"{Guid.NewGuid()}_{item.Name}.csv");
                    ToFile(item.Csv, fileName);
                    files.Add(fileName);
                });

                files.ForEach(file =>
                {
                    Assert.True(File.Exists(file));
                });
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (Directory.Exists(exportDirectory))
                {
                    Directory.Delete(exportDirectory, true);
                }
            }

        }

        private string BuildTempPath()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), $"OLT_UnitTest_EPPlus_{Guid.NewGuid()}");
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }
            return tempDir;
        }


        private void ToFile(byte[] stream, string saveToFileName)
        {

            if (File.Exists(saveToFileName))
            {
                File.Delete(saveToFileName);
            }

            var fileOutput = File.Create(saveToFileName, stream.Length);
            fileOutput.Write(stream, 0, stream.Length);
            fileOutput.Close();
        }
    }



}
