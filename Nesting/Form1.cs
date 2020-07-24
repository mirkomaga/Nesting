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
            lv.OwnerDraw = true;
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
    }
}
