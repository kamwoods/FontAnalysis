namespace WordDocAnalyzerAddin
{
    partial class DocumentAnalyzerForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Close = new System.Windows.Forms.Button();
            this.Search = new System.Windows.Forms.Button();
            this.Markup = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(376, 181);
            this.textBox1.TabIndex = 0;
            // 
            // Close
            // 
            this.Close.Location = new System.Drawing.Point(150, 370);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(75, 23);
            this.Close.TabIndex = 1;
            this.Close.Text = "Close";
            this.Close.UseVisualStyleBackColor = true;
            this.Close.Click += new System.EventHandler(this.Close_Click);
            // 
            // Search
            // 
            this.Search.Location = new System.Drawing.Point(136, 226);
            this.Search.Name = "Search";
            this.Search.Size = new System.Drawing.Size(108, 23);
            this.Search.TabIndex = 2;
            this.Search.Text = "Search Document";
            this.Search.UseVisualStyleBackColor = true;
            this.Search.Click += new System.EventHandler(this.Search_Click);
            // 
            // Markup
            // 
            this.Markup.Location = new System.Drawing.Point(135, 281);
            this.Markup.Name = "Markup";
            this.Markup.Size = new System.Drawing.Size(109, 23);
            this.Markup.TabIndex = 3;
            this.Markup.Text = "Markup Document";
            this.Markup.UseVisualStyleBackColor = true;
            this.Markup.Click += new System.EventHandler(this.Markup_Click);
            // 
            // DocumentAnalyzerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 405);
            this.Controls.Add(this.Markup);
            this.Controls.Add(this.Search);
            this.Controls.Add(this.Close);
            this.Controls.Add(this.textBox1);
            this.Name = "DocumentAnalyzerForm";
            this.Text = "Font Analyzer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Close;
        private System.Windows.Forms.Button Search;
        private System.Windows.Forms.Button Markup;

    }
}