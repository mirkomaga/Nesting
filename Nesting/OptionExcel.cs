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
        public static DataTable results;
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

            clbColonne.Items.Clear();
            clbColonne.Items.AddRange(columnNames);
        }
        private void btnConferma_Click(object sender, EventArgs e)
        {
            CheckedListBox.CheckedItemCollection checkedV = clbColonne.CheckedItems;
            IDictionary<string, dynamic> tmp = new Dictionary<string, dynamic>();

            DataSet ds = new DataSet("nesting");

            if (checkedV.Count > 0)
            {
                foreach (string columnName in checkedV)
                {
                    //tmp[columnName] = dt.DefaultView.ToTable(false, columnName).Rows;
                    DataTable table = dt.DefaultView.ToTable(false, columnName);
                    table.TableName = columnName;
                    ds.Tables.Add(table);
                }

                int lunghezzaT = ds.Tables[0].Rows.Count;
                int nTabella = ds.Tables.Count;

                results = new DataTable("tabellaJoinata");

                // popolo i nomi tabella

                for (int nt = 0; nt < nTabella; nt++)
                {
                    string tableName = ds.Tables[nt].TableName;
                    results.Columns.Add(tableName);
                }

                for (int lt = 0; lt < lunghezzaT; lt++)
                {
                    List<string> rws = new List<string>();

                    for (int nt = 0; nt < nTabella; nt++)
                    { 
                        var chilosa = ds.Tables[nt].Rows[lt][0].ToString();
                        rws.Add(chilosa);
                    }
                    results.Rows.Add(rws.ToArray());
                }
            }
            MessageBox.Show("Colonne caricate correttamente");
            this.Hide();
        }
        private void clbColonne_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void cbWorkSheet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
