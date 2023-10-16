namespace ExcelFileReader
{
    partial class DuplicateRowsForm
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
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.OKButton1 = new System.Windows.Forms.Button();
            this.ExportButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 11);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(1354, 518);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // OKButton1
            // 
            this.OKButton1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.OKButton1.Location = new System.Drawing.Point(692, 533);
            this.OKButton1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OKButton1.Name = "OKButton1";
            this.OKButton1.Size = new System.Drawing.Size(152, 38);
            this.OKButton1.TabIndex = 2;
            this.OKButton1.Text = "OK";
            this.OKButton1.UseVisualStyleBackColor = true;
            this.OKButton1.Click += new System.EventHandler(this.OKButton1_Click);
            // 
            // ExportButton
            // 
            this.ExportButton.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.ExportButton.Location = new System.Drawing.Point(541, 533);
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.Size = new System.Drawing.Size(145, 38);
            this.ExportButton.TabIndex = 3;
            this.ExportButton.Text = "EXPORT";
            this.ExportButton.UseVisualStyleBackColor = true;
            this.ExportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // DuplicateRowsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1378, 576);
            this.Controls.Add(this.ExportButton);
            this.Controls.Add(this.OKButton1);
            this.Controls.Add(this.richTextBox1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.Name = "DuplicateRowsForm";
            this.Text = "Duplicate Entries";
            this.Load += new System.EventHandler(this.DuplicateRowsForm_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button OKButton1;
        private System.Windows.Forms.Button ExportButton;
        //private System.Windows.Forms.DataGridView dataGridViewDuplicate;
    }
}