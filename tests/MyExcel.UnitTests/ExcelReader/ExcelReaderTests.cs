using System.IO;

namespace MyExcel.UnitTests
{
    public class ExcelReaderTests : IClassFixture<ExcelReaderTestsFixture>
    {
        private ExcelReaderTestsFixture _fixture;

        public ExcelReaderTests(ExcelReaderTestsFixture fixture)
        {
            _fixture = fixture;
        }

        [Theory]
        [InlineData("My Excel")]
        public void ShouldReadExcelFiles(string firstCellContent)
        {
            string actual;
            string expected = firstCellContent;

            ExcelReader reader = _fixture.GetAssginedExcelReader();

            actual = reader[1, 1];

            Assert.Equal(expected, actual);
        }

        [Fact]
        public void ReadingAnEmptyCellShouldReturnNull()
        {
            ExcelReader reader = _fixture.GetAssginedExcelReader();
            Assert.Null(reader[1, 10]);
        }

        [Fact]
        public void ReadingWithoutProvidingTheFileLocationThrowsAnException()
        {
            ExcelReader reader = _fixture.GetUnassignedExcelReader();
            Assert.Throws<FileLocationNotSetException>(() => reader[1, 1]);
        }

        [Fact]
        public void SettingAnInexistentFileLocationThrowsException()
        {
            ExcelReader reader = _fixture.GetUnassignedExcelReader();

            Assert.Throws<FileNotFoundException>(
                () => reader.FileLocation = _fixture.GetAnInexistentFileLocation());
        }

    }
}