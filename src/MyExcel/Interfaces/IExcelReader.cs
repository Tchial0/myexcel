using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace MyExcel
{
    /// <summary>
    /// Read excel files.
    /// </summary>
    public interface IExcelReader
    {

        /// <summary>
        /// Get/set the location of the excel file to be read.
        /// </summary>
        string FileLocation { get; set; }


        /// <summary>
        /// Read a cell value (no 0-based index).
        /// Throws an exception if the file location was not set.
        /// </summary>
        /// <param name="row">Index of the row in the spreadsheet</param>
        /// <param name="column">Index of the column in the apreadsheet</param>
        /// <returns>The value of the cell from the spreadsheet.</returns>
        string this[uint row, uint column] { get; }

        /// <summary>
        /// Get an enumerator of strings asynchronously representing a vertical selection from the spreadsheet.
        /// </summary>
        /// <param name="column">Index (no 0-based) of the column in the spreadsheet</param>
        /// <param name="startingRow">The row from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        Task<IEnumerable<string>> GetColumnAsync(uint column, uint startingRow = 1, CancellationToken cancellationToken = default);

        /// <summary>
        /// Get an enumerator of strings asynchronously representing an horizontal selection from the spreadsheet.
        /// </summary>
        /// <param name="row">Index (no 0-based) of the row in the spreadsheet</param>
        /// <param name="startingColumn">The column from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        Task<IEnumerable<string>> GetRowAsync(uint row, uint startingColumn = 1, CancellationToken cancellationToken = default);

        /// <summary>
        /// Get an enumerator of strings representing a vertical selection from the spreadsheet.
        /// </summary>
        /// <param name="column">Index (no 0-based) of the column in the spreadsheet</param>
        /// <param name="startingRow">The row from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        IEnumerable<string> GetColumn(uint column, uint startingRow = 1);

        /// <summary>
        /// Get an enumerator of strings representing an horizontal selection from the spreadsheet.
        /// </summary>
        /// <param name="row">Index (no 0-based) of the row in the spreadsheet</param>
        /// <param name="startingColumn">The column from which to start (default is 1).</param>
        /// <returns>The enumerator of strings representing the selection.</returns>
        IEnumerable<string> GetRow(uint row, uint startingColumn = 1);
    }
}
