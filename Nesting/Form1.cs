using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Nesting
{
    public partial class frm : Form
    {
        public static string pathExcel = null;
        public static string pathPart = null;
        public frm()
        {
            InitializeComponent();
        }

        private void frm_Load(object sender, EventArgs e)
        {
            lv.View = View.Details;
            lv.AllowColumnReorder = true;
            lv.FullRowSelect = true;
            lv.GridLines = true;
            lv.OwnerDraw = false;
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            frm.pathExcel = GenericFunction.chooseFile("Excel (*.xlsx)|*.xlsx");

            if (!string.IsNullOrEmpty(frm.pathExcel))
            {
                lblExcelD.Text = System.IO.Path.GetFileName(frm.pathExcel);
                Excel.analizzoExcel(frm.pathExcel);
            }
        }

        public void addTolv(string msg)
        {
            lv.Items.Add(msg);
        }

        private void btnInventor_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(frm.pathPart) && OptionExcel.results != null)
            {
                List<Rettangolo> oggetto = createObject();

                if (oggetto.Count > 0)
                {
                    InventorClass.CreateSM(oggetto, tspb, lv, frm.pathPart);
                }
            }
            else
            {
                MessageBox.Show("Scegliere la cartella di destinazione e selezionare l'Excel", "Attenzione");
            }
        }

        private static List<Rettangolo> createObject()
        {
            List<Rettangolo> results = new List<Rettangolo>();

            if (OptionExcel.results.Rows.Count > 0)
            {
                foreach( DataRow rw in OptionExcel.results.Rows)
                {
                    Rettangolo rett = new Rettangolo(rw, OptionExcel.results);
                    if (rett.status == true)
                    {
                        results.Add(rett);
                    }
                    else
                    {
                        MessageBox.Show("Non sono stati selezionati i dati necessari", "Attenzione");
                        break;
                    }
                }
            }

            return results;
        }

        private void toolStripProgressBar1_Click(object sender, EventArgs e)
        {

        }

        private void lv_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void lv_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {

            frm.pathPart = GenericFunction.chooseFolder();

            if (!string.IsNullOrEmpty(frm.pathPart))
            {
                lblInventorD.Text = System.IO.Path.GetFileName(frm.pathPart);
            }
        }
    }
}
