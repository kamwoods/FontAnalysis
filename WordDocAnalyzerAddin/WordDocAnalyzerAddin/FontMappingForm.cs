using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Word.Extensions;

using Map_Font;


namespace WordDocAnalyzerAddin
{

    
    [DllImport("MapFont.dll")]
    public static extern void map_font();

    public partial class FontMappingForm : Form
    {
        const int FONT_NOT_FOUND = -1;

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
        public FontMappingForm()
        {
            InitializeComponent();
/*
            this.fontDialog1.ShowDialog();

            Font font = this.fontDialog1.Font;
            System.Drawing.Graphics.AddFontResource()
            int indx;
            indx=0;

*/


            NativeMethods.map_font();



/*




            object oMissing = System.Reflection.Missing.Value;

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = Globals.ThisAddIn.Application;
            oDoc = oWord.ActiveDocument;

            // get all the paragraphs in the document

            ArrayList font_names = new ArrayList();
            string[] out_buf = new string[255];
            int indx = 0;


            int cnt, font_indx;

            Word.Range oRange;
            Word.Font oFont;

            // now do the document by char

            int num_chars, glyph_indx, num_fonts = 0;
            char cur_glyph;
            System.Collections.IEnumerator char_iter;

            out_buf[indx++] = "Font Mappings:";
            indx++;
            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;

            Graphics dc = this.CreateGraphics();
            System.IntPtr hDC = dc.GetHdc();
            System.Drawing.Font font;

            for (cnt = 0; cnt < num_chars; cnt++)
            {
                char_iter.MoveNext();
                oRange = (Microsoft.Office.Interop.Word.Range)char_iter.Current;
                oFont = oRange.Font;
                font_indx = search_font_table(font_names, oFont.Name);
                if (font_indx == FONT_NOT_FOUND)
                {
                    // add name to font table
                    font_names.Add(oFont.Name);
                    num_fonts++;
                    font_indx = num_fonts - 1;
                    // get mapping information
                    font = new Font((string)oFont.Name, 12);
                    out_buf[indx++] = "Font:    " + oFont.Name;
                    out_buf[indx++] = "Mapping: " + font.Name;
                    indx++;
                }
            }

            this.OutputBox.Lines = out_buf;
*/
        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
