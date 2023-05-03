﻿using System.Collections.Generic;

namespace MyExcel
{
    /// <summary>
    /// Write Excel files.
    /// </summary>
    public interface IExcelWriter : IExcel
    {
        string this[int row, int column] { set; }

        void WriteColumn(int column, IEnumerable<string> values, int startingRow = 1);

        void WriteRow(int row, IEnumerable<string> values, int startingColumn = 1);

        void SaveAs(string filename);
    }
}