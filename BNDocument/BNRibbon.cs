using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using System.Threading;
using BNDocument.Operations;
using BNDocument.Interfaces;

namespace BNDocument
{
    public partial class BNRibbon
    {
        int count = -1;
        private void BNRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var highlightWord = new HighlightWordOperation();
            var wordDoc = ICTDocument.GetICTDocument(highlightWord, "of");
            wordDoc.Execute();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var changeCaseUnderline = new UppercaseUnderlineOperation();
            changeCaseUnderline.Count = count++;
            var wordDoc = ICTDocument.GetICTDocument(changeCaseUnderline, "of");
            wordDoc.Execute();
        }


        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var reverseOperation = new ReverseOperation();
            var wordDoc = ICTDocument.GetICTDocument(reverseOperation, sender);
            wordDoc.Execute();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            var replaceOperation = new ReplaceMiddleWordOperation();
            var wordDoc = ICTDocument.GetICTDocument(replaceOperation, null);
            wordDoc.Execute();
        }
    }
}

