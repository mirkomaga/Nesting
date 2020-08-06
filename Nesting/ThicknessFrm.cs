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
    public partial class ThicknessFrm : Form
    {
        public static string pathPart = null;
        public ThicknessFrm()
        {
            InitializeComponent();
        }

        private void ThicknessFrm_Load(object sender, EventArgs e)
        {

        }

        private void btnFolder_Click(object sender, EventArgs e)
        {
            ThicknessFrm.pathPart = GenericFunction.chooseFolder(false);
            if (!string.IsNullOrEmpty(ThicknessFrm.pathPart))
            {
                tbInventorPart.Text = System.IO.Path.GetFileName(ThicknessFrm.pathPart);

                // TODO conto quanti ipt sono stati rilevati
                int iptCounter = GenericFunction.countFiles(ThicknessFrm.pathPart, "*.ipt");

                ListViewItem item1 = new ListViewItem("File ipt trovati: " + iptCounter.ToString(), 0);
                item1.SubItems.Add("Ok");
                lvThks.Items.AddRange(new ListViewItem[] { item1});
            }
        }

        private void btnAvvia_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(ThicknessFrm.pathPart))
            {
                InventorClass.changeThksSheetsMtl(ThicknessFrm.pathPart, lvThks, pbThk);
            }
            else
            {
                MessageBox.Show("Selezionare una cartella prima.", "Attenzione");
            }
        }
    }
}
