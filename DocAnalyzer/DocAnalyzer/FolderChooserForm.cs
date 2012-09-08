using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocAnalyzer
{
    public partial class FolderChooserForm : Form
    {
        private string _selectedFolder;

        public FolderChooserForm()
        {
            InitializeComponent();
            //this.folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.;
            this.folderBrowserDialog1.ShowDialog();
            this.SelectedFolder = this.folderBrowserDialog1.SelectedPath;
        }

        public string SelectedFolder
        {
            get { return _selectedFolder; }
            set { _selectedFolder = value; }
        }


        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }
    }
}
