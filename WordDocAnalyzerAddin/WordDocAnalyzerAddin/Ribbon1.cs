using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Word.Extensions;

namespace WordDocAnalyzerAddin
{
    public partial class Ribbon1 : OfficeRibbon
    {
        public Ribbon1()
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Analyze_Click(object sender, RibbonControlEventArgs e)
        {
            DocumentAnalyzerForm form = new DocumentAnalyzerForm();
            form.ShowDialog(); 
        }
/*
        private void Markup_Click(object sender, RibbonControlEventArgs e)
        {
            DocMarkupForm form = new DocMarkupForm();
            form.ShowDialog();
        }

        private void Mapping_Click(object sender, RibbonControlEventArgs e)
        {
            FontMappingForm form = new FontMappingForm();
            form.ShowDialog();
        }

        private void FontSearch_Click(object sender, RibbonControlEventArgs e)
        {
            DocSearchForm form = new DocSearchForm();
            form.ShowDialog();
        }
*/ 
   }


}
