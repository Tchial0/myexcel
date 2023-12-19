using System.Collections.Generic;
using System.Threading;
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
        /// Reads a cell value (no 0-based index).
        /// Throws an exception if the file location was not set.
        /// </summary>
        /// <param name="row">Index of the row in the spreadsheet</param>
        /// <param name="column">Index of the column in the apreadsheet</param>
        /// <returns>The value of the cell from the spreadsheet.</returns>
        public string this[uint row, uint column]
        {
            get
            {
                ThrowExceptionIfFileLocationNotSet();

                var cellValue = ((dynamic)_sheet.Cells[row, column]).Value;
       
                return (cellValue == null) ? null : cellValue.ToString();
            }
        }

        /// <summary>
        /// Gets/sets the location of the excel file to be read.
        /// When setting a file that does not exist an exception will be thrown.
        /// </summary>
        public string FileLocation
        {
            get { return _fileLocation; }
            set
            {
                if (System.IO.File.Exists(value) == false)
                {
                    throw new System.IO.FileNotFoundException($"Could not locate '{value}'");
                }

                _fileLocation = value;
                _workbook = _workbooks.Open(_fileLocation);
                _sheets = _workbook.Worksheets;
                _sheet = (Microsoft.Office.Interop.Excel.Worksheet)_sheets[1];
            }
        }

        /// <summary>
        /// Gets an enumerator of strings asynchronously representing a vertical selection from the spreadsheet.
        /// </summary>
        /// <param name="column">Index (no 0-based) of the column in the spreadsheet</param>
        /// <param name="startingRow">The row from which to start (default is 1).</param>
        /// <param name="cancellationToken">The token to cancel de task.</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public async Task<IEnumerable<string>> GetColumnAsync(uint column, uint startingRow = 1, CancellationToken cancellationToken = default)
        {
            ThrowExceptionIfFileLocationNotSet();

            return await Task.Run(() =>
            {
                List<string> values = new List<string>();
                for (uint row = startingRow; this[row, column] != string.Empty; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    values.Add(((dynamic)_sheet.Cells[row, column]).Value.ToString());
                }
                return values;
            });
        }

        /// <summary>
        /// Gets an enumerator of strings representing a vertical selection from the spreadsheet.
        /// </summary>
        /// <param name="column">Index (no 0-based) of the column in the spreadsheet</param>
        /// <param name="startingRow">The row from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public IEnumerable<string> GetColumn(uint column, uint startingRow = 1)
        {
            ThrowExceptionIfFileLocationNotSet();

            List<string> values = new List<string>();
            for (uint row = startingRow; this[row, column] != string.Empty; row++)
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
        /// <param name="cancellationToken">The token to cancel the task.</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public async Task<IEnumerable<string>> GetRowAsync(uint row, uint startingColumn = 1, CancellationToken cancellationToken = default )
        {
            ThrowExceptionIfFileLocationNotSet();

            return await Task.Run(() =>
            {
                List<string> values = new List<string>();
                for (uint column = startingColumn; this[row, column] != string.Empty; column++)
                {
                    values.Add(((dynamic)_sheet.Cells[row, column]).Value.ToString());
                }
                return values;
            });
        }

        /// <summary>
        /// Gets an enumerator of strings representing an horizontal selection from the spreadsheet.
        /// </summary>
        /// <param name="row">Index (no 0-based) of the row in the spreadsheet</param>
        /// <param name="startingColumn">The column from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        public IEnumerable<string> GetRow(uint row, uint startingColumn = 1)
        {
            ThrowExceptionIfFileLocationNotSet();

            List<string> values = new List<string>();
            for (uint column = startingColumn; this[row, column] != string.Empty; column++)
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
