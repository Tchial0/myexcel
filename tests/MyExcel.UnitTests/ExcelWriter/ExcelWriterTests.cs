using System.IO;

namespace MyExcel.UnitTests
{
    public class ExcelWriterTests  {

        [Fact]
        public void ExcelWriter_WhenSavingAFile_ShouldCreateTheFile()
        {
            var excelWriter = new ExcelWriter();
            var fileLocation = BaseFixture.GetAnInexistentFileLocation();

            excelWriter.SaveAs(fileLocation);
            excelWriter.Dispose();

            Assert.True(File.Exists(fileLocation));
            File.Delete(fileLocation);
        }
    }
}