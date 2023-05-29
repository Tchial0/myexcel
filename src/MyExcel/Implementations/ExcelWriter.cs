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
            _workbook = _workbooks.Add();
            _sheets = _workbook.Worksheets;
            _sheet = (Worksheet)_sheets.Add();
        }

        /// <summary>
        /// Writes a cell value (no 0-based index).
        /// </summary>
        /// <param name="row">Index of the row in the spreadsheet</param>
        /// <param name="column">Index of the column in the apreadsheet</param>
        public string this[uint row, uint column]
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
        public void WriteColumn(uint column, IEnumerable<string> values, uint startingRow = 1)
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
        public void WriteRow(uint row, IEnumerable<string> values, uint startingColumn = 1)
        {
            foreach (var value in values)
            {
                ((dynamic)_sheet.Cells[row, startingColumn++]).Value = value;
            }
        }

        /// <summary>
        /// Save the excel file.
        /// If another file with the same name exists will be deleted.
        /// </summary>
        /// <param name="filename">The full path of the the file including its extension (normally .xlsx).</param>
        public void SaveAs(string filename)
        {
            if (_sheet != null)
            {
                if (System.IO.File.Exists(filename))
                {
                    System.IO.File.Delete(filename);
                }
                _sheet.SaveAs(filename);
            }
        }
    }
}
