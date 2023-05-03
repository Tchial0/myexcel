using System.Collections.Generic;
using System.Threading.Tasks;

namespace MyExcel
{
    /// <summary>
    /// Read excel files.
    /// </summary>
    public interface IExcelReader
    {
       
        string FileLocation { get; set; }

        string this[int row, int column] { get; }

        Task<IEnumerable<string>> GetColumnAsync(int column, int startingRow = 1);

        Task<IEnumerable<string>> GetRowAsync(int row, int startingColumn = 1);

        IEnumerable<string> GetColumn(int column, int startingRow = 1);

        IEnumerable<string> GetRow(int row, int startingColumn = 1);
    }
}
