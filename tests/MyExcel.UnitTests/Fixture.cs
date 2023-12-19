namespace MyExcel.UnitTests
{
    public class Fixture : IDisposable
    {
        private readonly List<IDisposable> _elements;

        public Fixture()
        {
           _elements = new List<IDisposable>();
        }

        public IDisposable AddToDispose(IDisposable element)
        {
            _elements.Add(element);
            return element;
        }

        public void Dispose()
        {
           foreach(var element in _elements)
            {
                element.Dispose();
            }
        }
    }
}
