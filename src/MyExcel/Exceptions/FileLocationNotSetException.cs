using System;

namespace MyExcel
{
    /// <summary>
    /// Exception thrown when trying to access an inexistent file.
    /// </summary>
    public class FileLocationNotSetException : Exception
    {
        /// <summary>
        /// Initializes an instance of the FileLocationNotSetException class
        /// </summary>
        public FileLocationNotSetException() : base("The location of the Excel file was not set.")
        {
        }
    }
}
