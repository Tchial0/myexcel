namespace MyExcel.UnitTests
{
    public class ExcelReaderTests
    {

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