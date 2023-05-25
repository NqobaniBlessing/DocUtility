using BNDocument.Interfaces;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;

namespace BNDocument.Operations
{
    internal class ReverseOperation : IOperation
    {
        public void Execute(Document document, object context)
        {
            Range paragraphRange = Globals.ThisDocument.Application.Selection.Range;
            Range documentRange = document.Content;
            RibbonComboBox ribbonComboBox = (RibbonComboBox)context;

            // Get current selected value
            string currentText = ribbonComboBox.Text;

            if (currentText.Equals("Paragraph"))
            {
                // Get the paragraph range from it start to where the caret cursor is
                object start = paragraphRange.Start, end = paragraphRange.End;
                Range paragraphRange2 = document.Range(ref start, ref end);
                paragraphRange2.Expand(WdUnits.wdParagraph);

                // Split and reverse the current paragraph
                string text = paragraphRange2.Text;
                string[] words = text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                Array.Reverse(words);
                string reversedText = string.Join(" ", words);

                paragraphRange2.Text = reversedText;
            }
            else if (currentText.Equals("Document"))
            {
                // Get all the document paragraphs, and split and revers them
                string text = documentRange.Text;
                string[] paragraphs = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Array.Reverse(paragraphs);
                string reversedText = string.Join(Environment.NewLine, paragraphs);

                documentRange.Text = reversedText;
            }
        }
    }
}
