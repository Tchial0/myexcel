using System.Collections.Generic;
using System.Threading.Tasks;

namespace MyExcel
{
    /// <summary>
    /// Read excel files.
    /// </summary>
    public interface IExcelReader
    {
       
        void SetFileLocation(string path);

        string this[int row, int column] { get; }

        Task<IEnumerable<string>> GetColumnAsync(int column, int startingRow = 1);

        Task<IEnumerable<string>> GetRowAsync(int row, int startingColumn = 1);
    }
}
