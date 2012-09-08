namespace WordDocAnalyzerAddin
{
    partial class DocMarkupForm
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
            int indx, num_fonts, position ;

            this.Markup = new System.Windows.Forms.Button();
            this.Close = new System.Windows.Forms.Button();
            this.SuspendLayout();
            this.radio_buttons = new System.Windows.Forms.RadioButton[64];

            // 
            // radio_buttons
            // 


            position = 40;
            num_fonts = font_table.Count;

            for (indx = 0; indx < num_fonts; indx++)
            {
                this.radio_buttons[indx] = new System.Windows.Forms.RadioButton();
                this.radio_buttons[indx].AutoSize = true;
                this.radio_buttons[indx].Location = new System.Drawing.Point(40, position);
                this.radio_buttons[indx].Name = (string)font_table[indx];
                this.radio_buttons[indx].Size = new System.Drawing.Size(85, 20);
                this.radio_buttons[indx].TabIndex = 0;
                this.radio_buttons[indx].TabStop = true;
                this.radio_buttons[indx].Text = (string)font_table[indx];
                this.radio_buttons[indx].UseVisualStyleBackColor = true;
                position += 20;
            }
            this.radio_buttons[indx] = new System.Windows.Forms.RadioButton();
            this.radio_buttons[indx].AutoSize = true;
            this.radio_buttons[indx].Location = new System.Drawing.Point(40, position);
            this.radio_buttons[indx].Name = "!None!";
            this.radio_buttons[indx].Size = new System.Drawing.Size(85, 20);
            this.radio_buttons[indx].TabIndex = 0;
            this.radio_buttons[indx].TabStop = true;
            this.radio_buttons[indx].Text = "None";
            this.radio_buttons[indx].UseVisualStyleBackColor = true;
            position += 20;
            // 
            // Markup
            // 
            position += 20;
            this.Markup.Location = new System.Drawing.Point(80, position);
            this.Markup.Name = "Markup";
            this.Markup.Size = new System.Drawing.Size(116, 23);
            this.Markup.TabIndex = 1;
            this.Markup.Text = "Markup Document";
            this.Markup.UseVisualStyleBackColor = true;
            this.Markup.Click += new System.EventHandler(this.Markup_Click);
            // 
            // Close
            // 
            position += 40;
            this.Close.Location = new System.Drawing.Point(100, position);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(75, 23);
            this.Close.TabIndex = 2;
            this.Close.Text = "Close";
            this.Close.UseVisualStyleBackColor = true;
            this.Close.Click += new System.EventHandler(this.Close_Click);
            // 
            // Form1
            // 

            position += 100;
            if (position > 600)
                position = 600;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(275, position);
            this.Controls.Add(this.Close);
            this.Controls.Add(this.Markup);
            for (indx = 0; indx <= num_fonts; indx++)
            {
                this.Controls.Add(this.radio_buttons[indx]);
            }

            this.Name = "Font Markup";
            this.Text = "Font Markup";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton [] radio_buttons;
        private System.Windows.Forms.Button Markup;
        private System.Windows.Forms.Button Close;
    }
}