using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Add_txt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //Insert text in a range
            Word.Range rng = this.Application.ActiveDocument.Range(0, 0);
            rng.Text = "New Text";


            //Select the Range object, which has expanded from one character to the length of the inserted text.
            rng.Select();


            // Replace text in a Range
            Word.Range rng = this.Application.ActiveDocument.Range(0, 12);
            rng.Text = "New Text";


            // To insert text using the TypeText method
            //1. Declare a selection object variable
            Word.Selection currentSelection = Application.Selection;

            // Turn off the overtype if it is turn on
            if (Application.Options.Overtype)
            {
                Application.Options.Overtype = false;
            }
            // Test to see if selection is an insertion point.
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
            {
                currentSelection.TypeText("Inserting at insertion point. ");
                currentSelection.TypeParagraph();
            }
            else
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
            {
                // Move to start of selection.
                if (Application.Options.ReplaceSelection)
                {
                    object direction = Word.WdCollapseDirection.wdCollapseStart;
                    currentSelection.Collapse(ref direction);
                }
                currentSelection.TypeText("Inserting before a text block. ");
                currentSelection.TypeParagraph();
            }
            else
            {
                // Do nothing.
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0), 3, 4);
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorSeaGreen;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Font.Size = 12;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Rows.Borders.Enable = 1;
        }
    }
}
