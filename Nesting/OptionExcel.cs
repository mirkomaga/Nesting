using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Exc = Microsoft.Office.Interop.Excel;

namespace Nesting
{
    public partial class OptionExcel : Form
    {

        DataTableCollection dtbc;
        public static IDictionary<string, dynamic> results = new Dictionary<string, dynamic>();
        public OptionExcel(DataSet excO)
        {
            InitializeComponent();
            this.Show();

            this.popoloWs(excO);
        }
        private void popoloWs(DataSet excO)
        {
            cbWorkSheet.Items.Clear();
            dtbc = excO.Tables;
            cbWorkSheet.Items.AddRange(dtbc.Cast<DataTable>().Select(t => t.TableName).ToArray<string>());
        }
        private void OptionExcel_Load(object sender, EventArgs e)
        {
        }
        private void clbWorkSheets_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.popoloColonne(e.Index);
        }
        private void popoloColonne(int indexWs)
        {
            // DEVO TROVARE L'INTESTATURA
        }
        DataTable dt;
        private void cbWorkSheet_SelectionChangeCommitted(object sender, EventArgs e)
        {
            dt = dtbc[cbWorkSheet.SelectedItem.ToString()];

            var columnNames = (from c in dt.Columns.Cast<DataColumn>() select c.ColumnName).ToArray();

            clbColonne.Items.AddRange(columnNames);
        }
        private void btnConferma_Click(object sender, EventArgs e)
        {
            CheckedListBox.CheckedItemCollection checkedV = clbColonne.CheckedItems;

            if(checkedV.Count > 0)
            {
                foreach (string columnName in checkedV)
                {
                    results[columnName] = dt.DefaultView.ToTable(false, columnName);
                }
            }

            var result = MessageBox.Show("Colonne caricate correttamente");
            this.Hide();
        }
    }
}
