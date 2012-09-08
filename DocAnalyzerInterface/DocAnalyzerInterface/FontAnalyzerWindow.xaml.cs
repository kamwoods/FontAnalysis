using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Xml;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Word.Extensions;

namespace DocAnalyzerInterface
{
    /// <summary>
    /// Interaction logic for FontAnalyzerWindow.xaml
    /// </summary>
    public partial class FontAnalyzerWindow : Window
    {
        private Word._Application oWord;

        public FontAnalyzerWindow()
        {
            // start a version of Word to use for this app
            oWord = new Word.Application();
            oWord.Visible = false;

            InitializeComponent();
        }

        private void ProcessFile_Click(object sender, RoutedEventArgs e)
        {
            // choose a file to process
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\" ;
            openFileDialog1.Filter = "Word files (*.doc)|*.doc;*.docx" ;
            openFileDialog1.FilterIndex = 1 ;
            openFileDialog1.RestoreDirectory = true ;

            if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string doc_name ;

                // init global data
                font_name_tbl = new ArrayList();
                font_char_cnt = new int[MAX_NUM_FONTS];
                font_glyph_cnt = new int[MAX_NUM_FONTS];

                // process the file
                doc_name = openFileDialog1.FileName;
                AnalyzeDocument(doc_name);

                // output the result in XML format

                int indx, num_fonts;

                XmlTextWriter writer = new XmlTextWriter("c:\\fontinfo.xml", null);
                //Use indenting for readability.
                writer.Formatting = Formatting.Indented;

                writer.WriteComment("FontAnalysis");

                //Write an element (this one is the root).
                writer.WriteStartElement("FontAnalysis");

                //Write the namespace declaration.
                writer.WriteAttributeString("xmlns", "font", null, "urn:FontAnalysis");

                writer.WriteStartElement("file");

                string prefix = writer.LookupPrefix("urn:FontAnalysis");
                writer.WriteStartAttribute(prefix, "file_name", "urn:FontAnalysis");
                writer.WriteString(doc_name);
                writer.WriteEndAttribute();

                num_fonts = font_name_tbl.Count;
                
                //display information for each font in document
                for (indx = 0; indx < num_fonts; indx++)
                {
                    //Write the font.
                    writer.WriteStartElement("font");
                    writer.WriteStartAttribute(prefix, "name", "urn:FontAnalysis");
                    writer.WriteString((string)font_name_tbl[indx]);
                    writer.WriteEndAttribute();

                    //Write the char count.
                    writer.WriteStartElement("char");
                        writer.WriteStartElement("count");
                            writer.WriteString(font_char_cnt[indx].ToString());
                        writer.WriteEndElement();
                    writer.WriteEndElement();

                    //Write the glyph count.
                    writer.WriteStartElement("glyph");
                        writer.WriteStartElement("count");
                            writer.WriteString(font_glyph_cnt[indx].ToString());
                        writer.WriteEndElement();
                    writer.WriteEndElement();

                    //Write the end tag for the font element.
                    writer.WriteEndElement();
                }

                //Write the end tag for the file element.
                writer.WriteEndElement();

                //Write the end tag for the FontAnalysis element.
                writer.WriteEndElement();

                //Write the XML to file and close the writer.
                writer.Flush();
                writer.Close();
            }
        }

        private void BulkAnalyze_Click(object sender, RoutedEventArgs e)
        {
            // choose a directory

            // for each doc file in the directory 
            // process the file and output the result


        }

        const int MAX_NUM_FONTS = 255;
        const int MAX_NUM_GLYPHS = 255;
        const int FONT_NOT_FOUND = -1;
        const int GLYPH_NOT_FOUND = -1;

        private ArrayList font_name_tbl;
        private int[] font_char_cnt;
        private int[] font_glyph_cnt;

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

        public void AnalyzeDocument(string doc_name)
        {

            object oMissing = System.Reflection.Missing.Value;

            // master table for keeping track of which gyphs a font has used
            ArrayList font_glyphs = new ArrayList();

            // create a new document.
            Word._Document oDoc;

            object fileName = doc_name;
            object oReadOnly = (Boolean)true;
            object oVisible = (Boolean)false;

            oDoc = oWord.Documents.Open(ref fileName,
                ref oMissing, ref oReadOnly, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            int indx = 0;
            int cnt, font_indx;

            Word.Range oRange;
            Word.Font oFont;

            // now analyze the document by char

            int num_chars, glyph_indx, num_fonts=0;
            char cur_glyph;
            System.Collections.IEnumerator char_iter;

            char_iter = oDoc.Characters.GetEnumerator();
            num_chars = oDoc.Characters.Count;

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
                    glyph_indx = search_glyph_table( (char[])(font_glyphs[font_indx]), cur_glyph);
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


            return;

        }

    }

}
