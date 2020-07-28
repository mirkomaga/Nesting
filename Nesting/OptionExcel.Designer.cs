namespace Nesting
{
    partial class OptionExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OptionExcel));
            this.lblWorkSheets = new System.Windows.Forms.Label();
            this.cbWorkSheet = new System.Windows.Forms.ComboBox();
            this.clbColonne = new System.Windows.Forms.CheckedListBox();
            this.btnConferma = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblWorkSheets
            // 
            this.lblWorkSheets.AutoSize = true;
            this.lblWorkSheets.Location = new System.Drawing.Point(13, 27);
            this.lblWorkSheets.Name = "lblWorkSheets";
            this.lblWorkSheets.Size = new System.Drawing.Size(38, 13);
            this.lblWorkSheets.TabIndex = 2;
            this.lblWorkSheets.Text = "Sheet:";
            // 
            // cbWorkSheet
            // 
            this.cbWorkSheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbWorkSheet.FormattingEnabled = true;
            this.cbWorkSheet.Location = new System.Drawing.Point(57, 24);
            this.cbWorkSheet.Name = "cbWorkSheet";
            this.cbWorkSheet.Size = new System.Drawing.Size(238, 21);
            this.cbWorkSheet.TabIndex = 3;
            this.cbWorkSheet.SelectedIndexChanged += new System.EventHandler(this.cbWorkSheet_SelectedIndexChanged);
            this.cbWorkSheet.SelectionChangeCommitted += new System.EventHandler(this.cbWorkSheet_SelectionChangeCommitted);
            // 
            // clbColonne
            // 
            this.clbColonne.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.clbColonne.FormattingEnabled = true;
            this.clbColonne.Location = new System.Drawing.Point(16, 76);
            this.clbColonne.Name = "clbColonne";
            this.clbColonne.Size = new System.Drawing.Size(360, 214);
            this.clbColonne.TabIndex = 4;
            this.clbColonne.SelectedIndexChanged += new System.EventHandler(this.clbColonne_SelectedIndexChanged);
            // 
            // btnConferma
            // 
            this.btnConferma.Location = new System.Drawing.Point(301, 24);
            this.btnConferma.Name = "btnConferma";
            this.btnConferma.Size = new System.Drawing.Size(75, 23);
            this.btnConferma.TabIndex = 5;
            this.btnConferma.Text = "Coferma";
            this.btnConferma.UseVisualStyleBackColor = true;
            this.btnConferma.Click += new System.EventHandler(this.btnConferma_Click);
            // 
            // OptionExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(388, 325);
            this.Controls.Add(this.btnConferma);
            this.Controls.Add(this.clbColonne);
            this.Controls.Add(this.cbWorkSheet);
            this.Controls.Add(this.lblWorkSheets);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "OptionExcel";
            this.Text = "OptionExcel";
            this.Load += new System.EventHandler(this.OptionExcel_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblWorkSheets;
        private System.Windows.Forms.ComboBox cbWorkSheet;
        private System.Windows.Forms.CheckedListBox clbColonne;
        private System.Windows.Forms.Button btnConferma;
    }
}