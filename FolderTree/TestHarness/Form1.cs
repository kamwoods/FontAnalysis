using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestHarness
{
    public partial class Form1 : Form
    {
        private string selected_folder;

        public Form1()
        {
            InitializeComponent();
        }

        private void Select_Click(object sender, EventArgs e)
        {
            selected_folder = this.folderTree1.SelectedFolder;
            Console.WriteLine("you selected the folder: " + selected_folder);
            this.Close();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            selected_folder = "";
            this.Close();
        }
    }
}