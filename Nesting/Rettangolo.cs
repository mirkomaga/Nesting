using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Shapes;

namespace Nesting
{
    class Rettangolo
    {
        public int pezzi;
        public double lunghezza;
        public double sviluppo;
        public string commento;
        public string filename;
        public string marca;
        public bool status = true;

        public Rettangolo(DataRow dato, DataTable dt) 
        {
            var dict = dato.Table.Columns
              .Cast<DataColumn>()
              .ToDictionary(c => c.ColumnName, c => dato[c]);

            int index = dt.Rows.IndexOf(dato);

            //commenti = row+area
            if (dict.ContainsKey("AREA"))
            {
                filename = (index + 1).ToString() +"-"+ (string)dict["AREA"];
            }
            else
            {
                status = false;
            }

            //sviluppo = sv
            if (dict.ContainsKey("SV"))
            {
                sviluppo = (double)(Convert.ToDouble((string)dict["SV"]));
            }
            else
            {
                status = false;
            }

            //lunghezza = L di taglio
            if (dict.ContainsKey("L di taglio"))
            {
                lunghezza = (double)(Convert.ToDouble(dict["L di taglio"]));
            }
            else
            {
                status = false;
            }

            //marca = casing type horizontal
            if (dict.ContainsKey("CASING TYPE"))
            {
                marca = (string)dict["CASING TYPE"];
            }

            //pieces = nr
            if (dict.ContainsKey("NR"))
            {
                pezzi = (int)(Convert.ToInt64(dict["NR"]));
            }
        }
    }
}
