using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MyExcel
{
    /// <summary>
    /// Excel file reader.
    /// </summary>
    public class ExcelReader : Excel, IExcelReader
    {
        public string this[int row, int column]
        {
            get
            {
                string value;
                try
                {
                    value = ((dynamic)_sheet.Cells[row, column]).Value.ToString();
                
                }
                catch (Exception)
                {
                    value = string.Empty;
                }
                return value;
            }
        }

        public async Task<IEnumerable<string>> GetColumnAsync(int column, int startingRow = 1)
        {
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

        public async Task<IEnumerable<string>> GetRowAsync(int row, int startingColumn = 1)
        {
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

        public void SetFileLocation(string path)
        {
            if (_app != null)
            {
                _app.Workbooks.Close();
            }
            var wb = _app.Workbooks.Open(path);
            _sheet = (Worksheet)wb.Worksheets[1];
        }
    }
}
