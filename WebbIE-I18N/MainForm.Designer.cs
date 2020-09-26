namespace WebbIE_I18N
{
    partial class MainForm
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
            this.btnXMLToExcel = new System.Windows.Forms.Button();
            this.btnExcelToXML = new System.Windows.Forms.Button();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnXMLToExcel
            // 
            this.btnXMLToExcel.Location = new System.Drawing.Point(12, 12);
            this.btnXMLToExcel.Name = "btnXMLToExcel";
            this.btnXMLToExcel.Size = new System.Drawing.Size(167, 23);
            this.btnXMLToExcel.TabIndex = 0;
            this.btnXMLToExcel.Text = "WebbIE XML to Excel";
            this.btnXMLToExcel.UseVisualStyleBackColor = true;
            this.btnXMLToExcel.Click += new System.EventHandler(this.btnXMLToExcel_Click);
            // 
            // btnExcelToXML
            // 
            this.btnExcelToXML.Location = new System.Drawing.Point(12, 41);
            this.btnExcelToXML.Name = "btnExcelToXML";
            this.btnExcelToXML.Size = new System.Drawing.Size(167, 23);
            this.btnExcelToXML.TabIndex = 1;
            this.btnExcelToXML.Text = "Excel to WebbIE XML";
            this.btnExcelToXML.UseVisualStyleBackColor = true;
            this.btnExcelToXML.Click += new System.EventHandler(this.button2_Click);
            // 
            // lblLanguage
            // 
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(12, 67);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(55, 13);
            this.lblLanguage.TabIndex = 2;
            this.lblLanguage.Text = "Language";
            // 
            // cboLanguage
            // 
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Items.AddRange(new object[] {
            "de",
            "fr",
            "it",
            "sv"});
            this.cboLanguage.Location = new System.Drawing.Point(15, 83);
            this.cboLanguage.Name = "cboLanguage";
            this.cboLanguage.Size = new System.Drawing.Size(121, 21);
            this.cboLanguage.TabIndex = 3;
            this.cboLanguage.Text = "fr";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.cboLanguage);
            this.Controls.Add(this.lblLanguage);
            this.Controls.Add(this.btnExcelToXML);
            this.Controls.Add(this.btnXMLToExcel);
            this.Name = "MainForm";
            this.Text = "WebbIE I18N";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnXMLToExcel;
        private System.Windows.Forms.Button btnExcelToXML;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cboLanguage;
    }
}

