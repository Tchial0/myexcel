using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace MyExcel
{
    /// <summary>
    /// Excel file writer.
    /// </summary>
    public class ExcelWriter : Excel, IExcelWriter
    {
        /// <summary>
        /// Initializes a new instance of the ExcelWriter class.
        /// </summary>
        public ExcelWriter() : base()
        {
            Workbook wb = _app.Workbooks.Add();
            _sheet = (Worksheet)wb.Worksheets.Add();
        }

        /// <summary>
        /// Writes a cell value (no 0-based index).
        /// </summary>
        /// <param name="row">Index of the row in the spreadsheet</param>
        /// <param name="column">Index of the column in the apreadsheet</param>
        public string this[int row, int column]
        {
            set
            {
                ((dynamic)_sheet.Cells[row, column]).Value = value;
            }
        }

        /// <summary>
        /// Writes a vertical selection in the spreadsheet.
        /// </summary>
        /// <param name="column">The index (no 0-based) of the column.</param>
        /// <param name="values">The values to distribute across the selection</param>
        /// <param name="startingRow">The index (no 0-based) of the row from which to start.</param>
        public void WriteColumn(int column, IEnumerable<string> values, int startingRow = 1)
        {
            foreach (var value in values)
            {
                ((dynamic)_sheet.Cells[startingRow++, column]).Value = value;
            }
        }

        /// <summary>
        /// Writes an horizontal selection in the spreadsheet.
        /// </summary>
        /// <param name="row">The index (no 0-based) of the row.</param>
        /// <param name="values">The values to distribute across the selection</param>
        /// <param name="startingColumn">The index (no 0-based) of the column from which to start.</param>
        public void WriteRow(int row, IEnumerable<string> values, int startingColumn = 1)
        {
            foreach (var value in values)
            {
                ((dynamic)_sheet.Cells[row, startingColumn++]).Value = value;
            }
        }

        /// <summary>
        /// Save the excel file.
        /// </summary>
        /// <param name="filename">The name of the file</param>
        public void SaveAs(string filename)
        {
            if (_sheet != null)
            {
                _sheet.SaveAs(filename);
            }
        }
    }
}
