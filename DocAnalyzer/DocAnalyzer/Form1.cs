using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Reflection;

namespace DocAnalyzer
{
    public partial class Form1 : Form
    {
        private OpenFileDialog choose_file_dlg = new OpenFileDialog();
        private string chosen_file;




        public Form1()
        {
            InitializeComponent();
            choose_file_dlg.FileOk += new CancelEventHandler(choose_file_dlg_OK);
        }

        void choose_file_dlg_OK(object sender, CancelEventArgs e)
        {
            chosen_file = choose_file_dlg.FileName;
            this.Text = Path.GetFileName(chosen_file);
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

        private void button1_Click(object sender, EventArgs e)
        {
            // get name of file to process

            choose_file_dlg.ShowDialog();

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; // \endofdoc is a predefined bookmark

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;

            int str_indx;
            string out_file;

            str_indx = chosen_file.IndexOf('.');
            out_file = chosen_file.Insert(str_indx, ".fontinfo");
            StreamWriter sw = new StreamWriter(out_file);

            object fileName = chosen_file;
            object oReadOnly = (Boolean)true;

            oDoc = oWord.Documents.Open(ref fileName,
                ref oMissing, ref oReadOnly, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // get all the paragraphs in the document

            ArrayList font_name_tbl = new ArrayList(100);
            string[] out_buf = new string[255];
            int indx = 0;
            int num_para, cnt;
            Word.Paragraph oPara;
            Word.Range oRange;
            Word.Font oFont;

            /*            System.Collections.IEnumerator para_iter;

                        para_iter = oDoc.Paragraphs.GetEnumerator();
                        num_para = oDoc.Paragraphs.Count;
                        sw.WriteLine("number of paragraphs: " + num_para);
                        out_buf[indx++] = "number of paragraphs: " + Convert.ToString(num_para);

                        for (cnt = 0; cnt < num_para; cnt++)
                        {
                            para_iter.MoveNext();
                            oPara = (Microsoft.Office.Interop.Word.Paragraph)para_iter.Current;
                            oRange = oPara.Range;
                            oFont = oRange.Font;
                            //sw.WriteLine(oFont.NameAscii);
                            if (!search_array(font_name_tbl, oFont.NameAscii))
                                font_name_tbl.Add(oFont.NameAscii);
                        }

                        // now do the document by word

                        int num_words;
                        System.Collections.IEnumerator word_iter;

                        word_iter = oDoc.Words.GetEnumerator();
                        num_words = oDoc.Words.Count;
                        sw.WriteLine("number of words: " + num_words);
                        out_buf[indx++] = "number of words: " + Convert.ToString(num_words);

                        for (cnt = 0; cnt < num_words; cnt++)
                        {
                            word_iter.MoveNext();
                            oRange = (Microsoft.Office.Interop.Word.Range)word_iter.Current;
                            oFont = oRange.Font;
                            //sw.WriteLine(oFont.NameAscii);
                            if (!search_array(font_name_tbl, oFont.NameAscii))
                                font_name_tbl.Add(oFont.NameAscii);
                        }
*/
            // now do the document by char

            int num_chars;
            System.Collections.IEnumerator char_iter;

            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;
            sw.WriteLine("number of chars: " + num_chars);
            out_buf[indx++] = "number of characters: " + Convert.ToString(num_chars);



            for (cnt = 0; cnt < num_chars; cnt++)
            {
                char_iter.MoveNext();
                oRange = (Microsoft.Office.Interop.Word.Range)char_iter.Current;
                oFont = oRange.Font;
            /*    if (!search_array(font_name_tbl, oFont.NameAscii))
                {
                    font_name_tbl.Add(oFont.NameAscii);
                    sw.WriteLine(oFont.NameAscii);
                }
             */
                if (!search_array(font_name_tbl, oFont.Name))
                {
                    font_name_tbl.Add(oFont.Name);
                    sw.WriteLine(oFont.Name);
                }
            }

            
            object save_changes = Word.WdSaveOptions.wdDoNotSaveChanges;

            oWord.Documents.Close(ref save_changes, ref oMissing, ref oMissing);
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);


            sw.WriteLine("Dump of font name table:");
            indx++;
            out_buf[indx++] = "Dump of font name table:";
            foreach (string font in font_name_tbl)
            {
                sw.WriteLine(font);
                out_buf[indx++] = font;
            }

            this.textBox1.Lines = out_buf;

            sw.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
