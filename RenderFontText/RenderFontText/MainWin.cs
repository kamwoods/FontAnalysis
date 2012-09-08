using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace RenderFontText
{
    public partial class MainWin : Form
    {
        private Word.Application oWord;

        public MainWin()
        {
            oWord = new Word.Application();
            oWord.Visible = true;

            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_file_dlg = new OpenFileDialog();

            open_file_dlg.InitialDirectory = "c:\\";
            open_file_dlg.Filter = "Word 2003 Documents (*.doc)|*.doc|Word Documents (*.docx)|*.docx";
            open_file_dlg.FilterIndex = 2;
            open_file_dlg.RestoreDirectory = true;


            if (open_file_dlg.ShowDialog() == DialogResult.OK)
            {
                FontSelection font_select_form = new FontSelection(oWord, open_file_dlg.FileName);
                font_select_form.Show();
                //font_select_form.Dispose();
            }
        }
    }
}
