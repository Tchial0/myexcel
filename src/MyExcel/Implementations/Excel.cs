using Microsoft.Office.Interop.Excel;

namespace MyExcel
{

    /// <summary>
    /// The base class for excel readers and writers.
    /// </summary>    
    public abstract class Excel : IExcel
    {
        protected Application _app;
        protected Worksheet _sheet;

        protected Excel()
        {
            _app = new Application();

        }

        /// <summary>
        /// Releases all the resources used by the excel application.
        /// </summary>
        public void Dispose()
        {
            if (_app != null)
            {
                _app.Workbooks.Close();
                _app.Quit();
            }
        }
    }
}
