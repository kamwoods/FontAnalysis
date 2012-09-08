using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace BulkDocAnalyzer
{
    public partial class BulkDocAnalyzerForm : Form
    {
        private string selected_folder;
        private string[] out_buf = new string[128];
        private int buf_indx = 0;

        public BulkDocAnalyzerForm()
        {
            InitializeComponent();
            this.textBox1.Text = "Now we start to process the files...";
        }

        private void SelectFolder_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;

            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                selected_folder = this.folderBrowserDialog1.SelectedPath;

                Word._Application oWord;
                oWord = new Word.Application();
                oWord.Visible = false;

                if (selected_folder != "")
                {
                    string[] files = Directory.GetFiles(selected_folder, "*.doc");

                    foreach (string file in files)
                    {
                        out_buf[buf_indx++] = "Now analyzing: " + file;
                        this.textBox1.Lines = out_buf;
                        doc_analyzer(oWord, file);
                    }
                }

                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
        }

        const int MAX_NUM_FONTS = 255;
        const int MAX_NUM_GLYPHS = 255;
        const int FONT_NOT_FOUND = -1;
        const int GLYPH_NOT_FOUND = -1;

        private int search_font_table(ArrayList fonts, string font_name)
        {
            int indx, num_fonts;

            num_fonts = fonts.Count;
            for (indx = 0; indx < num_fonts; indx++)
            {
                if (((string)fonts[indx]).Equals(font_name))
                    return indx;
            }

            return FONT_NOT_FOUND;
        }

        private int search_glyph_table(char[] glyph_table, char glyph)
        {
            int indx, len;

            len = glyph_table.Length;
            for (indx = 0; indx < len; indx++)
            {
                if (glyph_table[indx] == glyph)
                    return indx;
            }

            return GLYPH_NOT_FOUND;
        }

        private void doc_analyzer(Word._Application oWord, string fname)
        {

            object oMissing = System.Reflection.Missing.Value;

            // master table for keeping track of which gyphs a font has used
            ArrayList font_glyphs = new ArrayList();
            int[] font_char_cnt = new int[MAX_NUM_FONTS];
            int[] font_glyph_cnt = new int[MAX_NUM_GLYPHS];

            //Start Word and create a new document.
            Word._Document oDoc;

            int str_indx;
            string out_file;

            str_indx = fname.IndexOf('.');
            out_file = fname.Substring(0, str_indx);
            out_file = out_file.Insert(str_indx,".fontinfo");
            StreamWriter sw = new StreamWriter(out_file);

            //out_buf[buf_indx++] = "Processing: " + fname;
            //this.textBox1.Lines = out_buf;

            object fileName = fname;
            object oReadOnly = (Boolean)true;
            object oVisible = (Boolean)false;

            oDoc = oWord.Documents.Open(ref fileName,
                ref oMissing, ref oReadOnly, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // get all the paragraphs in the document

            ArrayList font_name_tbl = new ArrayList(100);
            int cnt, font_indx;
            Word.Range oRange;
            Word.Font oFont;
            int num_chars, glyph_indx, num_fonts = 0;
            char cur_glyph;
            System.Collections.IEnumerator char_iter;

            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;

            sw.WriteLine("Number of Chars: " + num_chars);
            sw.WriteLine();

            // now do the document by char

            for (cnt = 0; cnt < num_chars; cnt++)
            {
                char_iter.MoveNext();
                oRange = (Microsoft.Office.Interop.Word.Range)char_iter.Current;
                oFont = oRange.Font;
                font_indx = search_font_table(font_name_tbl, oFont.Name);
                if (font_indx == FONT_NOT_FOUND)
                {
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
            sw.WriteLine();
            sw.WriteLine();

            //object save_changes = Word.WdSaveOptions.;

            //oWord.Documents.Close(ref oMissing, ref oMissing, ref oMissing);

            sw.WriteLine("Dump of font name table and font usage information:");
            sw.WriteLine();
            sw.WriteLine("\tName\t\tNumber of Chars\t\t\tNumber of Glyphs");

            font_indx = 0;
            foreach (string font in font_name_tbl)
            {
                sw.WriteLine(font + "\t\t\t" + font_char_cnt[font_indx] + "\t\t\t\t" + font_glyph_cnt[font_indx]);
                font_indx++;
            }

            sw.Close();        
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
