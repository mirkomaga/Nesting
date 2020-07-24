using Exc = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;
using ExcelDataReader;
using System.Data;

namespace Nesting
{
    class Excel
    {
        public static DataTableCollection dtbc;
        public static void analizzoExcel(string pathExcel)
        {
            try
            {
                using (var stream = File.Open(@pathExcel, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        manageOptionsExcel(result);
                    }
                }
            }
            catch (System.IO.IOException)
            {

            }

            //selezionare foglio da analizzare
        }
        private static void manageOptionsExcel(DataSet excO)
        {
            new OptionExcel(excO);
        }
    }
}
