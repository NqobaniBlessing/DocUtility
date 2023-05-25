using BNDocument.Interfaces;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BNDocument.Operations
{
    internal class ReplaceMiddleWordOperation : IOperation
    {
        public void Execute(Document document, object context)
        {
            int paragraphCount = document.Paragraphs.Count;
            for (int i = 2; i <= paragraphCount; i += 2)
            {
                if (i % 2 == 0)
                {
                    Paragraph previousParagraph = document.Paragraphs[i - 1];

                    string middleWord = GetMiddleWord(previousParagraph.Range.Text);

                    Paragraph currentParagraph = document.Paragraphs[i];

                    ReplaceFirstWord(currentParagraph.Range, middleWord);
                }
            }
        }

        // Helper methods
        private string GetMiddleWord(string text)
        {
            string[] words = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            int middleIndex = words.Length / 2;

            return words[middleIndex];
        }

        private void ReplaceFirstWord(Range range, string replacement)
        {
            Find find = range.Find;
            find.ClearFormatting();
            find.Text = GetMiddleWord(range.Text);
            find.MatchWildcards = true;
            bool found = find.Execute();

            if (found)
            {
                range.Text = replacement + range.Text.Substring(find.Text.Length);
            }
        }
    }
}
