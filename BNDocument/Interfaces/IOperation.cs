using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BNDocument.Interfaces
{
    internal interface IOperation
    {
        void Execute(Document document, object context);
    }
}
