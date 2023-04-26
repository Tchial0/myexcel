using System;
using System.Collections.Generic;
using System.Text;

namespace MyExcel
{
    /// <summary>
    /// Exception thrown when trying to access an inexistent file.
    /// </summary>
    public class FileLocationNotSetException : Exception
    {
        /// <summary>
        /// 
        /// </summary>
        public FileLocationNotSetException() : base("The location of the Excel file was not set.")
        {
            
        }
    }
}
