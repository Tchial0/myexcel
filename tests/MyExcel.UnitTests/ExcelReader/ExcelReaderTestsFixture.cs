using System.IO;

namespace MyExcel.UnitTests
{
    public class ExcelReaderTestsFixture : BaseFixture
    {
        public ExcelReader GetUnassignedExcelReader()
        {
            var excelReader = new ExcelReader();
            AddToDispose(excelReader);
            return excelReader;
        }

        public ExcelReader GetAssginedExcelReader(){
            var excelReader = GetUnassignedExcelReader();
            var basePath = Directory.GetCurrentDirectory();

            excelReader.FileLocation = Path.Combine(
                basePath,
                ExcelReaderTestsConstants.LocalExcelFileName);
                
            return excelReader;
        }

        public string GetAnInexistentFileLocation()
        {
            var tempName = Path.GetTempFileName();
            File.Delete(tempName);
            return tempName;
        }

    }
}