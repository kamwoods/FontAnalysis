using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LOGFontGetter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.fontDialog1.ShowDialog();
            object tmp = this.fontDialog1;
        
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
