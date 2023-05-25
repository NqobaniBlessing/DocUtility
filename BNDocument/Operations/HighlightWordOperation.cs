using BNDocument.Interfaces;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BNDocument.Operations
{
    internal class HighlightWordOperation : IOperation
    {
        public void Execute(Document document, object context)
        {
            int count = 0;
            var word = (string)context;

            foreach (Range range in document.Words)
            {
                if (range.Text.Trim().Equals(word, StringComparison.OrdinalIgnoreCase))
                {
                    range.HighlightColorIndex = WdColorIndex.wdYellow;
                    count++;
                }
            }

            MessageBox.Show($"The word '{word}' occurred {count} times",
                "Word Highlighter", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
