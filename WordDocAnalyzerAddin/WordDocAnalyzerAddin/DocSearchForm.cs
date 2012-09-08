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
    public partial class DocSearchForm : Form
    {

        public DocSearchForm()
        {
            make_font_table();
            InitializeComponent();
        }

        private ArrayList font_table = new ArrayList();
        private string search_font;

        private Boolean search_array(ArrayList fonts, string font_name)
        {
            foreach (string font in fonts)
            {
                if (font.Equals(font_name))
                    return true;
            }

            return false;
        }

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

        private void next_occur(Word._Document oDoc, string search_font, ref int start_pos, ref int end_pos)
        {

            Word.Range oRange;
            Word.Font oFont;
            int length, pos;
            object position;

            length = oDoc.Characters.Count;
            if (end_pos == (length-1))
                pos = 0;
            else
                pos = end_pos;

            // find the beginning of the next occurance of the search font

            while (true)
            {
                position = pos;
                oRange = oDoc.Range(ref position, ref position);
                oFont = oRange.Font;
                if (oFont.Name.Equals(search_font))
                   break;
                else
                    pos++;
                if (pos == length)
                    pos = 0;
            }

            start_pos = pos;

            // find the end of this occurance

            while (true)
            {
                position = pos;
                oRange = oDoc.Range(ref position, ref position);
                oFont = oRange.Font;
                if (!oFont.Name.Equals(search_font))
                    break;
                else
                    pos++;
                if (pos == length)
                {
                    pos--;
                    break;
                }
            }

            end_pos = pos;

            return;

        }

        private int start_pos = 0, end_pos = 0;

        private void Search_Click(object sender, EventArgs e)
        {
            // get which font has been selected for marking

            int indx;

            for (indx=0; indx <= font_table.Count; indx++)
            {
                if (radio_buttons[indx].Checked)
                {
                    search_font = radio_buttons[indx].Name;
                    break;
                }
            }

            // get the next occurance of this font in the document

            Word._Application oWord;
            Word._Document oDoc;
            Word.Range oRange;
            object start, end;

            oWord = Globals.ThisAddIn.Application;
            oDoc = oWord.ActiveDocument; 

            next_occur(oDoc, search_font, ref start_pos, ref end_pos);

            // and mark the range as selected

            start = start_pos;
            end = end_pos;
            oRange = oDoc.Range(ref start, ref end);
            oRange.Select();

        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
