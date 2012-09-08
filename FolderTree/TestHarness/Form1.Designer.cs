namespace TestHarness
{
    partial class Form1
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
            this.folderTree1 = new FolderTree.FolderTree();
            this.Select = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.folderTree1)).BeginInit();
            this.SuspendLayout();
            // 
            // folderTree1
            // 
            this.folderTree1.Location = new System.Drawing.Point(6, 3);
            this.folderTree1.Name = "folderTree1";
            this.folderTree1.RootFolder = "c:\\Program Files\\7-Zip";
            this.folderTree1.Size = new System.Drawing.Size(287, 355);
            this.folderTree1.TabIndex = 0;
            this.folderTree1.ShowFiles = false;
            // 
            // Select
            // 
            this.Select.Location = new System.Drawing.Point(90, 375);
            this.Select.Name = "Select";
            this.Select.Size = new System.Drawing.Size(109, 23);
            this.Select.TabIndex = 1;
            this.Select.Text = "Select Folder";
            this.Select.UseVisualStyleBackColor = true;
            this.Select.Click += new System.EventHandler(this.Select_Click);
            // 
            // Cancel
            // 
            this.Cancel.Location = new System.Drawing.Point(105, 413);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(75, 23);
            this.Cancel.TabIndex = 2;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(298, 439);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.Select);
            this.Controls.Add(this.folderTree1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.folderTree1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private FolderTree.FolderTree folderTree1;
        private System.Windows.Forms.Button Select;
        private System.Windows.Forms.Button Cancel;
    }
}

