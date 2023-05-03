using System.Collections.Generic;
using System.Threading.Tasks;

namespace MyExcel
{
    /// <summary>
    /// Excel file reader.
    /// </summary>
    public class ExcelReader : Excel, IExcelReader
    {
        private string _fileLocation = null;

        /// <summary>
        /// Read a cell value (no 0-based index).
        /// Throws an exception if the file location was not set.
        /// </summary>
        /// <param name="row">Index of the row in the spreadsheet</param>
        /// <param name="column">Index of the column in the apreadsheet</param>
        /// <returns>The value of the cell from the spreadsheet.</returns>
        public string this[int row, int column]
        {
            get
            {
                ThrowExceptionIfFileLocationNotSet();

                return ((dynamic)_sheet.Cells[row, column]).Value.ToString();
            }
        }

        /// <summary>
        /// Get/set the location of the excel file to be read.
        /// </summary>
        public string FileLocation
        {
            get { return _fileLocation; }
            set
            {
                if (_app != null)
                {
                    _app.Workbooks.Close();
                }

                _fileLocation = value;
                var wb = _app.Workbooks.Open(_fileLocation);
                _sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            }
        }

        /// <summary>
        /// Get an enumerator of strings asynchronously representing a vertical selection from the spreadsheet.
        /// </summary>
        /// <param name="column">Index (no 0-based) of the column in the spreadsheet</param>
        /// <param name="startingRow">The row from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public async Task<IEnumerable<string>> GetColumnAsync(int column, int startingRow = 1)
        {
            ThrowExceptionIfFileLocationNotSet();

            return await Task.Run(() =>
            {
                List<string> values = new List<string>();
                for (int row = startingRow; this[row, column] != string.Empty; row++)
                {
                    values.Add(((dynamic)_sheet.Cells[row, column]).Value.ToString());
                }
                return values;
            });
        }

        /// <summary>
        /// Get an enumerator of strings asynchronously representing a vertical selection from the spreadsheet.
        /// </summary>
        /// <param name="column">Index (no 0-based) of the column in the spreadsheet</param>
        /// <param name="startingRow">The row from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public IEnumerable<string> GetColumn(int column, int startingRow = 1)
        {
            ThrowExceptionIfFileLocationNotSet();

            List<string> values = new List<string>();
            for (int row = startingRow; this[row, column] != string.Empty; row++)
            {
                values.Add(((dynamic)_sheet.Cells[row, column]).Value.ToString());
            }
            return values;
        }

        /// <summary>
        /// Get an enumerator of strings asynchronously representing an horizontal selection from the spreadsheet.
        /// </summary>
        /// <param name="row">Index (no 0-based) of the row in the spreadsheet</param>
        /// <param name="startingColumn">The column from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public async Task<IEnumerable<string>> GetRowAsync(int row, int startingColumn = 1)
        {
            ThrowExceptionIfFileLocationNotSet();

            return await Task.Run(() =>
            {
                List<string> values = new List<string>();
                for (int column = startingColumn; this[row, column] != string.Empty; column++)
                {
                    values.Add(((dynamic)_sheet.Cells[row, column]).Value.ToString());
                }
                return values;
            });
        }

        /// <summary>
        /// Get an enumerator of strings representing an horizontal selection from the spreadsheet.
        /// </summary>
        /// <param name="row">Index (no 0-based) of the row in the spreadsheet</param>
        /// <param name="startingColumn">The column from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public IEnumerable<string> GetRow(int row, int startingColumn = 1)
        {
            ThrowExceptionIfFileLocationNotSet();

            List<string> values = new List<string>();
            for (int column = startingColumn; this[row, column] != string.Empty; column++)
            {
                values.Add(((dynamic)_sheet.Cells[row, column]).Value.ToString());
            }
            return values;
        }

        private void ThrowExceptionIfFileLocationNotSet()
        {
            if (string.IsNullOrEmpty(_fileLocation)) throw new FileLocationNotSetException();
        }

    }


}
