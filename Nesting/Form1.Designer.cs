namespace Nesting
{
    partial class frm
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.boxExcel = new System.Windows.Forms.GroupBox();
            this.lblExcelD = new System.Windows.Forms.Label();
            this.lblExcelS = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.stsBott = new System.Windows.Forms.StatusStrip();
            this.tspb = new System.Windows.Forms.ToolStripProgressBar();
            this.lv = new System.Windows.Forms.ListView();
            this.boxInventor = new System.Windows.Forms.GroupBox();
            this.lblInventorD = new System.Windows.Forms.Label();
            this.btnInventor = new System.Windows.Forms.Button();
            this.lblInvetorS = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.chOperazione = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chStato = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.boxExcel.SuspendLayout();
            this.stsBott.SuspendLayout();
            this.boxInventor.SuspendLayout();
            this.SuspendLayout();
            // 
            // boxExcel
            // 
            this.boxExcel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.boxExcel.Controls.Add(this.lblExcelD);
            this.boxExcel.Controls.Add(this.lblExcelS);
            this.boxExcel.Controls.Add(this.btnExcel);
            this.boxExcel.Location = new System.Drawing.Point(12, 12);
            this.boxExcel.Name = "boxExcel";
            this.boxExcel.Size = new System.Drawing.Size(339, 99);
            this.boxExcel.TabIndex = 0;
            this.boxExcel.TabStop = false;
            this.boxExcel.Text = "Excel";
            // 
            // lblExcelD
            // 
            this.lblExcelD.AutoSize = true;
            this.lblExcelD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExcelD.Location = new System.Drawing.Point(43, 44);
            this.lblExcelD.Name = "lblExcelD";
            this.lblExcelD.Size = new System.Drawing.Size(115, 13);
            this.lblExcelD.TabIndex = 2;
            this.lblExcelD.Text = "Nessun file selezionato";
            // 
            // lblExcelS
            // 
            this.lblExcelS.AutoSize = true;
            this.lblExcelS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExcelS.Location = new System.Drawing.Point(6, 44);
            this.lblExcelS.Name = "lblExcelS";
            this.lblExcelS.Size = new System.Drawing.Size(31, 13);
            this.lblExcelS.TabIndex = 1;
            this.lblExcelS.Text = "File:";
            // 
            // btnExcel
            // 
            this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExcel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btnExcel.Location = new System.Drawing.Point(258, 39);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 23);
            this.btnExcel.TabIndex = 0;
            this.btnExcel.Text = "Sfoglia";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // stsBott
            // 
            this.stsBott.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tspb});
            this.stsBott.Location = new System.Drawing.Point(0, 480);
            this.stsBott.Name = "stsBott";
            this.stsBott.Size = new System.Drawing.Size(363, 22);
            this.stsBott.TabIndex = 1;
            this.stsBott.Text = "statusStrip1";
            // 
            // tspb
            // 
            this.tspb.Name = "tspb";
            this.tspb.Size = new System.Drawing.Size(100, 16);
            this.tspb.Click += new System.EventHandler(this.toolStripProgressBar1_Click);
            // 
            // lv
            // 
            this.lv.AllowDrop = true;
            this.lv.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chOperazione,
            this.chStato});
            this.lv.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lv.FullRowSelect = true;
            this.lv.GridLines = true;
            this.lv.HideSelection = false;
            this.lv.Location = new System.Drawing.Point(0, 285);
            this.lv.Name = "lv";
            this.lv.ShowItemToolTips = true;
            this.lv.Size = new System.Drawing.Size(363, 195);
            this.lv.TabIndex = 2;
            this.lv.UseCompatibleStateImageBehavior = false;
            this.lv.View = System.Windows.Forms.View.Details;
            this.lv.DrawColumnHeader += new System.Windows.Forms.DrawListViewColumnHeaderEventHandler(this.lv_DrawColumnHeader);
            this.lv.DrawSubItem += new System.Windows.Forms.DrawListViewSubItemEventHandler(this.lv_DrawSubItem);
            // 
            // boxInventor
            // 
            this.boxInventor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.boxInventor.Controls.Add(this.lblInventorD);
            this.boxInventor.Controls.Add(this.btnInventor);
            this.boxInventor.Controls.Add(this.lblInvetorS);
            this.boxInventor.Controls.Add(this.button1);
            this.boxInventor.Location = new System.Drawing.Point(12, 131);
            this.boxInventor.Name = "boxInventor";
            this.boxInventor.Size = new System.Drawing.Size(339, 136);
            this.boxInventor.TabIndex = 3;
            this.boxInventor.TabStop = false;
            this.boxInventor.Text = "Inventor";
            // 
            // lblInventorD
            // 
            this.lblInventorD.AutoSize = true;
            this.lblInventorD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInventorD.Location = new System.Drawing.Point(66, 43);
            this.lblInventorD.Name = "lblInventorD";
            this.lblInventorD.Size = new System.Drawing.Size(142, 13);
            this.lblInventorD.TabIndex = 5;
            this.lblInventorD.Text = "Nessuna cartella selezionata";
            // 
            // btnInventor
            // 
            this.btnInventor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInventor.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btnInventor.Location = new System.Drawing.Point(258, 93);
            this.btnInventor.Name = "btnInventor";
            this.btnInventor.Size = new System.Drawing.Size(75, 23);
            this.btnInventor.TabIndex = 0;
            this.btnInventor.Text = "Avvia";
            this.btnInventor.UseVisualStyleBackColor = false;
            this.btnInventor.Click += new System.EventHandler(this.btnInventor_Click);
            // 
            // lblInvetorS
            // 
            this.lblInvetorS.AutoSize = true;
            this.lblInvetorS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInvetorS.Location = new System.Drawing.Point(6, 43);
            this.lblInvetorS.Name = "lblInvetorS";
            this.lblInvetorS.Size = new System.Drawing.Size(54, 13);
            this.lblInvetorS.TabIndex = 4;
            this.lblInvetorS.Text = "Cartella:";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.button1.Location = new System.Drawing.Point(258, 38);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "Sfoglia";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // chOperazione
            // 
            this.chOperazione.Text = "Operazione";
            this.chOperazione.Width = 299;
            // 
            // chStato
            // 
            this.chStato.Text = "Stato";
            // 
            // frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 502);
            this.Controls.Add(this.boxInventor);
            this.Controls.Add(this.lv);
            this.Controls.Add(this.stsBott);
            this.Controls.Add(this.boxExcel);
            this.Name = "frm";
            this.Text = "Nesting from Excel";
            this.Load += new System.EventHandler(this.frm_Load);
            this.boxExcel.ResumeLayout(false);
            this.boxExcel.PerformLayout();
            this.stsBott.ResumeLayout(false);
            this.stsBott.PerformLayout();
            this.boxInventor.ResumeLayout(false);
            this.boxInventor.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox boxExcel;
        private System.Windows.Forms.Label lblExcelD;
        private System.Windows.Forms.Label lblExcelS;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.StatusStrip stsBott;
        private System.Windows.Forms.GroupBox boxInventor;
        private System.Windows.Forms.Button btnInventor;
        private System.Windows.Forms.Label lblInventorD;
        private System.Windows.Forms.Label lblInvetorS;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.ListView lv;
        public System.Windows.Forms.ToolStripProgressBar tspb;
        private System.Windows.Forms.ColumnHeader chOperazione;
        private System.Windows.Forms.ColumnHeader chStato;
    }
}

