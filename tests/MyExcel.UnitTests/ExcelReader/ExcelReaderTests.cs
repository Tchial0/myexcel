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

        [Fact]
        public void ExcelReader_WhenReadingNonEmptyCell_ShouldReturnTheRealValue()
        {
            ExcelReader reader = _fixture.GetAssginedExcelReader();
            string expected = ExcelReaderTestsConstants.LocalExcelFirstCellContent;
            string actual = reader[1, 1];

            Assert.Equal(expected, actual);
        }

        [Fact]
        public void ExcelReader_WhenReadingAnEmptyCell_ShouldReturnNull()
        {
            ExcelReader reader = _fixture.GetAssginedExcelReader();
            Assert.Null(reader[1, 10]);
        }

        [Fact]
        public void ExcelReader_WhenReadingWithoutProvidingTheFileLocation_ShouldThrowAnException()
        {
            ExcelReader reader = _fixture.GetUnassignedExcelReader();
            Assert.Throws<FileLocationNotSetException>(() => reader[1, 1]);
        }

        [Fact]
        public void ExcelReader_WhenSettingAnInexistentFileLocation_ShouldThrowAnException()
        {
            ExcelReader reader = _fixture.GetUnassignedExcelReader();

            Assert.Throws<FileNotFoundException>(
                () => reader.FileLocation = _fixture.GetAnInexistentFileLocation());
        }

    }
}