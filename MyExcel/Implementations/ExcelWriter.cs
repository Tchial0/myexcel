using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace MyExcel
{
    /// <summary>
    /// Excel file writer.
    /// </summary>
    public class ExcelWriter : Excel, IExcelWriter
    {
        public ExcelWriter() : base()
        {
            Workbook wb = _app.Workbooks.Add();
            _sheet = (Worksheet)wb.Worksheets.Add();
        }

        public string this[int row, int column]
        {
            set
            {
                ((dynamic)_sheet.Cells[row, column]).Value = value;
            }
        }

        public void WriteColumn(int column, IEnumerable<string> values, int startingRow = 1)
        {
            foreach (var value in values)
            {
                ((dynamic)_sheet.Cells[startingRow++, column]).Value = value;
            }
        }

        public void WriteRow(int row, IEnumerable<string> values, int startingColumn = 1)
        {
            foreach (var value in values)
            {
                ((dynamic)_sheet.Cells[row, startingColumn++]).Value = value;
            }
        }

        public void SaveAs(string filename)
        {
            if (_sheet != null)
            {
                _sheet.SaveAs(filename);
            }
        }
    }
}
