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

        public string this[int row, int column]
        {
            get
            {
                ThrowExceptionIfFileLocationNotSet();

                return ((dynamic)_sheet.Cells[row, column]).Value.ToString();
            }
        }

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
