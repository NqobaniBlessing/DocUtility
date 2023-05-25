using BNDocument.Interfaces;
using Microsoft.Office.Interop.Word;
using System;
using System.Text;

namespace BNDocument.Operations
{
    internal class UppercaseUnderlineOperation : IOperation
    {
        private int _count;
        public int Count { get => _count; set => _count = value; }

        public void Execute(Document document, object context)
        {
            var word = (string)context;
            var stringBuilder = new StringBuilder();
            Range range = document.Content;
            Style style = document.Styles["Normal"];

            stringBuilder.Append(range.Text);

            Find contextWordFind = document.Content.Find;
            contextWordFind.Text = word;

            if (_count % 2 != 0)
            {
                style.Font.Underline = WdUnderline.wdUnderlineNone;
                contextWordFind.Replacement.Text = word.ToUpper();
                contextWordFind.MatchCase = false;
                contextWordFind.MatchWholeWord = true;
                contextWordFind.Execute(Replace: WdReplace.wdReplaceAll);
            }
            else
            {
                range.Text = stringBuilder.Replace(word.ToUpper(), word.ToLower()).ToString();
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                Find precedingWordFind = document.Content.Find;
                precedingWordFind.Text = word;
                precedingWordFind.Forward = true;

                bool found = precedingWordFind.Execute();

                if (found)
                {
                    Range foundRange = contextWordFind.Parent;

                    for (int i = 1; i < foundRange.Words.Count; i++)
                    {
                        if (foundRange.Words[i].Text.Trim().Equals("of", StringComparison.Ordinal))
                        {
                            foundRange.Words[i - 1].Font.Underline = WdUnderline.wdUnderlineSingle;
                        }
                    }
                }
            }
            document.Fields.Update();
        }
    }
}
