namespace RenderFontText
{
    partial class FontSelection
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
            int indx, num_fonts, position;

            this.RenderText = new System.Windows.Forms.Button();
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
            // 
            // RenderText
            // 
            position += 20;
            this.RenderText.Location = new System.Drawing.Point(87, position);
            this.RenderText.Name = "RenderText";
            this.RenderText.Size = new System.Drawing.Size(105, 23);
            this.RenderText.TabIndex = 0;
            this.RenderText.Text = "Render Text";
            this.RenderText.UseVisualStyleBackColor = true;
            this.RenderText.Click += new System.EventHandler(this.RenderText_Click);

            // 
            // Close
            // 
            position += 40;
            this.Close.Location = new System.Drawing.Point(104, position);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(75, 23);
            this.Close.TabIndex = 1;
            this.Close.Text = "Close";
            this.Close.UseVisualStyleBackColor = true;
            this.Close.Click += new System.EventHandler(this.Close_Click);

            // 
            // FontSelection
            // 

            position += 100;
            if (position > 600)
                position = 600;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(275, position);
            this.Controls.Add(this.Close);
            this.Controls.Add(this.RenderText);
            for (indx = 0; indx<num_fonts; indx++)
            {
                this.Controls.Add(this.radio_buttons[indx]);
            }

            this.Name = "RenderFontText";
            this.Text = "Render Text of Selected Font";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.RadioButton[] radio_buttons;
        private System.Windows.Forms.Button RenderText;
        private System.Windows.Forms.Button Close;
    }
}