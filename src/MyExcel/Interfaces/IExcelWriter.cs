using System.Collections.Generic;

namespace MyExcel
{
    /// <summary>
    /// Write Excel files.
    /// </summary>
    public interface IExcelWriter : IExcel
    {
        /// <summary>
        /// Writes a cell value (no 0-based index).
        /// </summary>
        /// <param name="row">Index of the row in the spreadsheet</param>
        /// <param name="column">Index of the column in the apreadsheet</param>
        string this[uint row, uint column] { set; }


        /// <summary>
        /// Writes a vertical selection in the spreadsheet.
        /// </summary>
        /// <param name="column">The index (no 0-based) of the column.</param>
        /// <param name="values">The values to distribute across the selection</param>
        /// <param name="startingRow">The index (no 0-based) of the row from which to start.</param>
        void WriteColumn(uint column, IEnumerable<string> values, uint startingRow = 1);

        /// <summary>
        /// Writes an horizontal selection in the spreadsheet.
        /// </summary>
        /// <param name="row">The index (no 0-based) of the row.</param>
        /// <param name="values">The values to distribute across the selection</param>
        /// <param name="startingColumn">The index (no 0-based) of the column from which to start.</param>
        void WriteRow(uint row, IEnumerable<string> values, uint startingColumn = 1);


        /// <summary>
        /// Save the excel file.
        /// If another file with the same name exists will be deleted.
        /// </summary>
        /// <param name="filename">The full path of the the file including its extension (normally .xlsx).</param>
        void SaveAs(string filename);
    }
}
