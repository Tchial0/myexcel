using System.Diagnostics;

namespace MyExcel.UnitTests
{
    public class ExcelReaderTests
    {
        
        [Theory]
        [InlineData(@"C:\Users\HP\source\repos\myexcel\lib\MyExcel.xlsx", "MyExcel")]
        public void ShouldReadExcelFiles(string excelFileLocation, string firstCellContent)
        {
            
            string actual;
            string expected = firstCellContent;

            using (ExcelReader reader = new ExcelReader())
            {
                reader.FileLocation = excelFileLocation;
                actual = reader[1, 1];
            }
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void ReadingWithoutProvidingTheFileLocationThrowsAnException()
        {
            ExcelReader reader = new ExcelReader();
            bool threwException = false;

            try
            {
                var cell = reader[1, 1];
            }
            catch (FileLocationNotSetException)
            {
                threwException = true;
            }
            finally
            {
                reader.Dispose();
            }

            Assert.True(threwException);
        }

        [Fact]
        public void SettingAnInexistentFileThrowsAnException()
        {
            ExcelReader reader = new ExcelReader();
            bool threwException = false;

            try
            {
                reader.FileLocation = GetAnInexistentFileLocation();
            }
            catch (FileNotFoundException)
            {
                threwException = true;
            }
            finally
            {
                reader.Dispose();
            }

            Assert.True(threwException);
        }

        private string GetAnInexistentFileLocation()
        {
            var tempName = System.IO.Path.GetTempFileName();
            System.IO.File.Delete(tempName);
            return tempName;
        }
    }
}