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
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Word.Extensions;

using Map_Font;


namespace WordDocAnalyzerAddin
{
    public partial class DocumentAnalyzerForm : Form
    {
        const int MAX_NUM_FONTS = 255;
        const int MAX_NUM_GLYPHS = 255;
        const int FONT_NOT_FOUND = -1;
        const int GLYPH_NOT_FOUND = -1;

        private ArrayList font_name_tbl;

        private int search_font_table(ArrayList fonts, string font_name)
        {
            int indx, num_fonts ;

            num_fonts = fonts.Count;
            for( indx=0; indx<num_fonts; indx++)
            {
                if (((string)fonts[indx]).Equals(font_name))
                    return indx;
            }

            return FONT_NOT_FOUND;
        }

        private int search_glyph_table(char[] glyph_table, char glyph)
        {
            int indx, len ;

            len = glyph_table.Length;
            for( indx=0; indx<len; indx++)
            {
                if(glyph_table[indx] == glyph)
                    return indx;
            }

            return GLYPH_NOT_FOUND;
        }

        public DocumentAnalyzerForm()
        {
            InitializeComponent();

            object oMissing = System.Reflection.Missing.Value;

            // master able for keeping track of which gyphs a font has used
            ArrayList font_glyphs = new ArrayList();
            int[] font_char_cnt = new int[MAX_NUM_FONTS];
            int[] font_glyph_cnt = new int[MAX_NUM_FONTS];

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = Globals.ThisAddIn.Application ;
            oDoc = oWord.ActiveDocument;

            // get all the paragraphs in the document

            font_name_tbl = new ArrayList();
            string[] out_buf = new string[255];
            int indx = 0;


            int cnt, font_indx;

            Word.Range oRange;
            Word.Font oFont;

            // now do the document by char

            int num_chars, glyph_indx, num_fonts=0;
            char cur_glyph;
            System.Collections.IEnumerator char_iter;

            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;
            out_buf[indx++] = "number of characters: " + Convert.ToString(num_chars);

            for (cnt = 0; cnt < num_chars; cnt++)
            {
                char_iter.MoveNext();
                oRange = (Microsoft.Office.Interop.Word.Range)char_iter.Current;
                oFont = oRange.Font;
                font_indx = search_font_table(font_name_tbl, oFont.Name);
                if (font_indx == FONT_NOT_FOUND)
                {
                    // add name to font table
                    font_name_tbl.Add(oFont.Name);
                    num_fonts++;
                    font_indx = num_fonts - 1;
                    // create new font glyph list and add current glyph
                    font_glyphs.Add(new char[MAX_NUM_GLYPHS]); 
                    ((char [])font_glyphs[font_indx])[0] = oRange.Text[0] ;
                    // start char count for this font
                    font_char_cnt[font_indx] = 1;
                    font_glyph_cnt[font_indx] = 1;
                }else{
                    cur_glyph = oRange.Text[0] ;
                    glyph_indx = search_glyph_table( (char[])font_glyphs[font_indx], cur_glyph);
                    if (glyph_indx == GLYPH_NOT_FOUND)
                    {
                        font_glyph_cnt[font_indx]++;
                        glyph_indx = font_glyph_cnt[font_indx] - 1;
                        ((char[])font_glyphs[font_indx])[glyph_indx] = cur_glyph;
                    }
                    // start char count for this font
                    font_char_cnt[font_indx]++; 
                }

           }

            indx++;
            out_buf[indx++] = "Dump of font name table:";
            font_indx = 0;
            foreach (string font in font_name_tbl)
            {
                out_buf[indx++] = "Name: " + font + "   number of chars: " + font_char_cnt[font_indx] +
                                   "   number of glyphs: " + font_glyph_cnt[font_indx];
                font_indx++;
            }

            this.textBox1.Lines = out_buf;
        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Search_Click(object sender, EventArgs e)
        {
            DocSearchForm form = new DocSearchForm();
            form.ShowDialog();
        }

        private void Markup_Click(object sender, EventArgs e)
        {
            DocMarkupForm form = new DocMarkupForm();
            form.ShowDialog();
        }
    }
}
