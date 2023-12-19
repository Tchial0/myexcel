using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace MyExcel
{
    /// <summary>
    /// The base class for excel readers and writers.
    /// </summary>    
    public abstract class Excel : IExcel
    {
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
        protected Application _app;
        protected Worksheet _sheet;
        protected Workbooks _workbooks;
        protected Workbook _workbook;
        protected Sheets _sheets;
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member

        /// <summary>
        /// Initializes a new instance of a Excel file reader or writer.
        /// </summary>    
        protected Excel()
        {
            _app = new Application();
            _workbooks = _app.Workbooks;
            _app.DisplayAlerts = false;
        }

        /// <summary>
        /// Releases all the resources used by the excel application.
        /// </summary>
        public void Dispose()
        {
            if (_app != null)
            {
                if (_sheet != null)
                {
                    Marshal.FinalReleaseComObject(_sheet);
                    _sheet = null;
                }

                if (_sheets != null)
                {
                    Marshal.FinalReleaseComObject(_sheets);
                    _sheets = null;
                }

                if (_workbook != null)
                {
                    _workbook.Close(false);
                    Marshal.FinalReleaseComObject(_workbook);
                    _workbook = null;
                }

                if (_workbooks != null)
                {
                    _workbooks.Close();
                    Marshal.FinalReleaseComObject(_workbooks);
                    _workbooks = null;
                }

                _app.Quit();
                Marshal.FinalReleaseComObject(_app);
                _app = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

            }
        }
    }
}
