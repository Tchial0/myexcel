namespace MyExcel.UnitTests
{
    public class ExcelReaderTests
    {

        [Fact]
        public void ReadingWithoutProvidingTheFileLocationThrowsAnException()
        {
            MyExcel.ExcelReader reader = new MyExcel.ExcelReader();
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

    }
}