using System;
using System.Collections.Generic;
using System.Text;

namespace MyExcel
{
    public class FileLocationNotSetException : Exception
    {
        public FileLocationNotSetException() : base("The location of the Excel file was not set.")
        {
            
        }
    }
}
