namespace MyExcel.UnitTests
{
    public class ExcelReaderTests : IClassFixture<Fixture>
    {
        private Fixture _fixture;

        public ExcelReaderTests(Fixture fixture)
        {
            _fixture = fixture;
        }

        [Theory]
        [InlineData(@"C:\Users\HP\source\repos\myexcel\lib\MyExcel.xlsx", "My Excel")]
        public void ShouldReadExcelFiles(string excelFileLocation, string firstCellContent)
        {
            string actual;
            string expected = firstCellContent;

            ExcelReader reader = GetExcelReader();

            reader.FileLocation = excelFileLocation;
            actual = reader[1, 1];

            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData(@"C:\Users\HP\source\repos\myexcel\lib\MyExcel.xlsx")]
        public void ReadingInexistentCellShouldReturnNull(string excelFileLocation)
        {
            ExcelReader reader = GetExcelReader();

            reader.FileLocation = excelFileLocation;

            Assert.Null(reader[1, 10]);
        }

        [Fact]
        public void ReadingWithoutProvidingTheFileLocationThrowsAnException()
        {
            ExcelReader reader = GetExcelReader();
            Assert.Throws<FileLocationNotSetException>(() => reader[1, 1]);
        }

        [Fact]
        public void SettingAnInexistentFileThrowsAnException()
        {
            ExcelReader reader = GetExcelReader();
            Assert.Throws<FileNotFoundException>(() => reader.FileLocation = GetAnInexistentFileLocation());
           
        }

        private string GetAnInexistentFileLocation()
        {
            var tempName = System.IO.Path.GetTempFileName();
            System.IO.File.Delete(tempName);
            return tempName;
        }

        private ExcelReader GetExcelReader() => (ExcelReader)_fixture.AddToDispose(new ExcelReader());

    }
}