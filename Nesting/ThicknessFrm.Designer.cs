namespace Nesting
{
    partial class ThicknessFrm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbInventorPart = new System.Windows.Forms.TextBox();
            this.btnFolder = new System.Windows.Forms.Button();
            this.lvThks = new System.Windows.Forms.ListView();
            this.chOperazione = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chStato = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnAvvia = new System.Windows.Forms.Button();
            this.pbThk = new System.Windows.Forms.ProgressBar();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.btnAvvia);
            this.groupBox1.Controls.Add(this.btnFolder);
            this.groupBox1.Controls.Add(this.tbInventorPart);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(301, 139);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Inventor Part";
            // 
            // tbInventorPart
            // 
            this.tbInventorPart.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbInventorPart.Location = new System.Drawing.Point(6, 37);
            this.tbInventorPart.Name = "tbInventorPart";
            this.tbInventorPart.Size = new System.Drawing.Size(208, 20);
            this.tbInventorPart.TabIndex = 0;
            this.tbInventorPart.Text = "Nessuna cartella selezionata";
            // 
            // btnFolder
            // 
            this.btnFolder.Location = new System.Drawing.Point(222, 36);
            this.btnFolder.Name = "btnFolder";
            this.btnFolder.Size = new System.Drawing.Size(75, 23);
            this.btnFolder.TabIndex = 1;
            this.btnFolder.Text = "Sfoglia";
            this.btnFolder.UseVisualStyleBackColor = true;
            this.btnFolder.Click += new System.EventHandler(this.btnFolder_Click);
            // 
            // lvThks
            // 
            this.lvThks.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chOperazione,
            this.chStato});
            this.lvThks.GridLines = true;
            this.lvThks.HideSelection = false;
            this.lvThks.Location = new System.Drawing.Point(12, 157);
            this.lvThks.Name = "lvThks";
            this.lvThks.Size = new System.Drawing.Size(303, 304);
            this.lvThks.TabIndex = 7;
            this.lvThks.UseCompatibleStateImageBehavior = false;
            this.lvThks.View = System.Windows.Forms.View.Details;
            // 
            // chOperazione
            // 
            this.chOperazione.Text = "Operazione";
            this.chOperazione.Width = 236;
            // 
            // chStato
            // 
            this.chStato.Text = "Stato";
            // 
            // btnAvvia
            // 
            this.btnAvvia.Location = new System.Drawing.Point(222, 89);
            this.btnAvvia.Name = "btnAvvia";
            this.btnAvvia.Size = new System.Drawing.Size(75, 23);
            this.btnAvvia.TabIndex = 2;
            this.btnAvvia.Text = "Avvia";
            this.btnAvvia.UseVisualStyleBackColor = true;
            this.btnAvvia.Click += new System.EventHandler(this.btnAvvia_Click);
            // 
            // pbThk
            // 
            this.pbThk.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pbThk.Location = new System.Drawing.Point(0, 470);
            this.pbThk.Name = "pbThk";
            this.pbThk.Size = new System.Drawing.Size(325, 23);
            this.pbThk.TabIndex = 8;
            // 
            // ThicknessFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(325, 493);
            this.Controls.Add(this.pbThk);
            this.Controls.Add(this.lvThks);
            this.Controls.Add(this.groupBox1);
            this.Name = "ThicknessFrm";
            this.Text = "ThicknessFrm";
            this.Load += new System.EventHandler(this.ThicknessFrm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnFolder;
        private System.Windows.Forms.TextBox tbInventorPart;
        private System.Windows.Forms.Button btnAvvia;
        private System.Windows.Forms.ColumnHeader chOperazione;
        private System.Windows.Forms.ColumnHeader chStato;
        public System.Windows.Forms.ListView lvThks;
        private System.Windows.Forms.ProgressBar pbThk;
    }
}