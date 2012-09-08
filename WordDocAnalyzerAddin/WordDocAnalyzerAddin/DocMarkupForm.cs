using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Reflection;

namespace WordDocAnalyzerAddin
{
    public partial class DocMarkupForm : Form
    {
        private ArrayList font_table = new ArrayList(100);
        private string markup_font;

        void make_font_table()
        {
            Word._Application oWord;
            Word._Document oDoc;
            int cnt;
            Word.Range oRange;
            Word.Font oFont;

            oWord = Globals.ThisAddIn.Application;
            oDoc = oWord.ActiveDocument;

            // now do the document by char

            int num_chars;
            System.Collections.IEnumerator char_iter;

            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;

            for (cnt = 0; cnt < num_chars; cnt++)
            {
                char_iter.MoveNext();
                oRange = (Microsoft.Office.Interop.Word.Range)char_iter.Current;
                oFont = oRange.Font;
                if (!search_array(font_table, oFont.Name))
                {
                    font_table.Add(oFont.Name);
                }
            }

        }

        public DocMarkupForm()
        {
            make_font_table();
            InitializeComponent();
        }

        private Boolean search_array(ArrayList fonts, string font_name)
        {
            foreach (string font in fonts)
            {
                if (font.Equals(font_name))
                    return true;
            }

            return false;
        }

        private void Markup_Click(object sender, EventArgs e)
        {

            // which font has been selected for marking

            int indx;

            for (indx=0; indx <= font_table.Count; indx++)
            {
                if (radio_buttons[indx].Checked)
                {
                    markup_font = radio_buttons[indx].Name;
                    break;
                }
            }

            Word._Application oWord;
            Word._Document oDoc;
            int cnt;
            Word.Range oRange;
            Word.Font oFont;

            oWord = Globals.ThisAddIn.Application;
            oDoc = oWord.ActiveDocument; 

            // now do the document by char

            int num_chars;
            System.Collections.IEnumerator char_iter;

            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;
           
            for (cnt = 0; cnt < num_chars; cnt++)
            {
                char_iter.MoveNext();
                oRange = (Microsoft.Office.Interop.Word.Range)char_iter.Current;
                oFont = oRange.Font;
                if (oFont.Name.Equals(markup_font))
                {
                    oRange.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdYellow;
                    
                }else{
                    oRange.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight;
                }


            }

        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
