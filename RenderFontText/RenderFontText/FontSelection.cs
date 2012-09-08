using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace RenderFontText
{
    public partial class FontSelection : Form
    {
        private Word._Application oWord;
        private Word._Document oDoc;
        private string fname = "";

        public FontSelection(Word._Application WordApp, string fname)
        {
            object oMissing = System.Reflection.Missing.Value;

            oWord = WordApp;
            object fileName = fname;
            object oReadOnly = (Boolean)true;
            object oVisible = (Boolean)true;

            oDoc = oWord.Documents.Open(ref fileName,
                ref oMissing, ref oReadOnly, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            make_font_table(fname);
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

        void make_font_table(string fname)
        {

            int cnt;
            Word.Range oRange;
            Word.Font oFont;


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

        private Boolean next_occur(string search_font, ref int start_pos, ref int end_pos)
        {

            Word.Range oRange;
            Word.Font oFont;
            int length, pos;
            object position;

            length = oDoc.Characters.Count;
            if (end_pos == (length - 1))
                //end of file
                return false;
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
                    // reached end of file
                    return false;
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

            return true;

        }

        private int start_pos, end_pos;

        private void RenderText_Click(object sender, EventArgs e)
        {
            // get which font has been selected for marking
            int indx;
            string render_font="";

            for (indx = 0; indx <= font_table.Count; indx++)
            {
                if (radio_buttons[indx].Checked)
                {
                    render_font = radio_buttons[indx].Name;
                    break;
                }
            }

            start_pos = 0;
            end_pos = 0;

            //FontTextView font_text_view_form = new FontTextView(render_font);
            // open new Word document for text rendering
            object file_name = @"c:\TestDocs\output.rnd";
            object newTemplate = false;
            object docType = 0;
            object isVisible = true;
            object oMissing = System.Reflection.Missing.Value;
            Word._Document oDoc_out;

            oDoc_out = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            
            // get the next occurance of this font in the document

            Word.Range oRange;
            object start, end;
            Word.Selection sln = oWord.Selection;

            while (next_occur(render_font, ref start_pos, ref end_pos))
            {
                start = start_pos;
                end = end_pos;
                oRange = oDoc.Range(ref start, ref end);
                sln.TypeText(oRange.Text);
                sln.TypeParagraph();
                //font_text_view_form.render_text(oRange.Text);
            }
            //oDoc_out.Close(ref oMissing, ref oMissing, ref oMissing);
            //font_text_view_form.Show();
        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
