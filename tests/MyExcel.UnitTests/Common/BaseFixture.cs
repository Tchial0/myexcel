using System;
using System.Collections.Generic;
using System.IO;

namespace MyExcel.UnitTests
{
    public abstract class BaseFixture : IDisposable
    {
        private readonly List<IDisposable> _elements;

        public BaseFixture()
        {
            _elements = new List<IDisposable>();
        }

        protected IDisposable AddToDispose(IDisposable element)
        {
            _elements.Add(element);
            return element;
        }

        public void Dispose()
        {
            foreach (var element in _elements)
            {
                element.Dispose();
            }
        }

        public static string GetAnInexistentFileLocation()
        {
            var tempName = Path.GetTempFileName();
            File.Delete(tempName);
            return tempName;
        }
    }
}
