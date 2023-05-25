using BNDocument.Interfaces;
using Microsoft.Office.Interop.Word;

namespace BNDocument
{
    internal class ICTDocument
    {
        private readonly Document _document = Globals.ThisDocument.Application.ActiveDocument;
        private IOperation _operation;
        private object _context;
        private static ICTDocument _ictDocument = new ICTDocument();

        internal IOperation Operation { get => _operation; set => _operation = value; }
        internal object Context { get => _context; set => _context = value; }

        private ICTDocument()
        {
        }

        // Ensure there is only one instance of the ICTDocument object througout the program's lifecycle
        public static ICTDocument GetICTDocument(IOperation operation, object context)
        {
            _ictDocument.Operation = operation;
            _ictDocument.Context = context;

            return _ictDocument;
        }
        public void Execute()
        {
            _operation.Execute(_document, _context);
        }

    }
}
