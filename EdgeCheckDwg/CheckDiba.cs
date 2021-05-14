using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EdgeAutocadPlugins
{
    public partial class CheckDiba : Form
    {
        public static string codiceCommessa = "200000";

        public CheckDiba()
        {
            InitializeComponent();
            this.listView1.AllowDrop = true;
            this.listView1.DragEnter += new DragEventHandler(lv_DragEnter);
            this.listView1.DragDrop += new DragEventHandler(lv_DragDrop);
        }

        private void CheckDiba_Load(object sender, EventArgs e)
        {

        }

        void lv_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void lv_DragDrop(object sender, DragEventArgs e)
        {
            // listView1.Items.Clear();

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files) {
                if (Path.GetExtension(file) == ".dwg")
                {
                    string[] row = { file };
                    var listViewItem = new ListViewItem(row);
                    listView1.Items.Add(listViewItem);

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count == 0) return;

            List<string> listaFiles = new List<string>();

            foreach (ListViewItem i in listView1.Items)
            {
                listaFiles.Add( i.Text );
            }

            // devo ciclare tutti i dwg, prendere diba e raccogliere codice ce

            List<ceC> lcec = Commands.CicloCE(listaFiles.ToArray());

            List<string> codiciCellula = lcec.Select(x => x.NOME).ToList();

            Distribuzione db = new Distribuzione();

            string q = "SELECT* FROM bim_200000.objects;";

            DataTable dtO = db.call(q);
            DataTable dtE = db.call("SELECT ele_name, ele_destination FROM bim_200000.elements;");

            Dictionary<string, int> assocCEID = new Dictionary<string, int>();

            foreach (string codeC in codiciCellula)
            {
                q = "SELECT ID FROM bim_200000.objects where obj_codice like '" + codeC + "%';";
                DataTable dt = db.call(q);

                int id = int.Parse(dt.Rows[0]["ID"].ToString());

                assocCEID.Add(codeC, id);
            }

            List<DistintaBase> dwgDIBA = new List<DistintaBase>();
            List<DistintaBase> dbDIBA = new List<DistintaBase>();

            foreach (KeyValuePair<string, int> c in assocCEID)
            {
                dwgDIBA = new List<DistintaBase>();
                dbDIBA = new List<DistintaBase>();

                // QUI CARICO SOLTANTO GLI ACCESSORI DAL DATABASE

                q = "SELECT * FROM bim_200000.obj_distinta where od_parentid = "+ c.Value +";";

                DataTable dt = db.call(q);

                ceC cellula = lcec.Find(x => x.NOME.Contains(c.Key));

                foreach (DataRow r in dt.Rows)
                {
                    DistintaBase diba = new DistintaBase();
                    
                    diba.TIPOLOGIA = c.Value.ToString();
                    diba.CODICE = r[2].ToString();
                    diba.QTA = r[4].ToString();
                    diba.UM = r[5].ToString();
                    diba.OC = r[6].ToString();
                    diba.DESCRIZIONE = r[3].ToString();
                    diba.SOURCE = "";
                    
                    dbDIBA.Add(diba);
                }

                // QUI CARICO ANCHE LE DISTINTE DAL DB

                DataRow[] elemento = dtO.Select("ID = '" + c.Value + "'");

                string compCodice = elemento[0]["obj_comp_codici"].ToString();

                foreach (string r in compCodice.Split('|'))
                {
                    if (string.IsNullOrWhiteSpace(r))
                        continue;

                    DistintaBase diba = new DistintaBase();

                    DataRow el = dtO.Select("obj_codice = '" + r + "'").First();

                    string innerType = el[3].ToString();

                    DataRow element = dtE.Select("ele_name = '" + innerType.ToString()+"'").First();

                    if (element["ele_destination"].ToString() == "0")
                    {
                        diba.TIPOLOGIA = c.Value.ToString();
                        diba.CODICE = codiceCommessa + r.ToString();
                        diba.QTA = "1";
                        diba.UM = "NR";
                        diba.OC = "O";
                        diba.DESCRIZIONE = el[4].ToString();
                        diba.SOURCE = "";

                        dbDIBA.Add(diba);
                    }

                }

                // QUI CARICO TUTTO DAL DWG

                foreach (DataRow r in cellula.DIBA.Rows)
                {
                    DistintaBase diba = new DistintaBase();

                    diba.TIPOLOGIA = r[0].ToString();
                    diba.CODICE = r[1].ToString();
                    diba.QTA = r[2].ToString();
                    diba.UM = r[3].ToString();
                    diba.OC = r[4].ToString();
                    diba.DESCRIZIONE = r[5].ToString();
                    diba.SOURCE = "";

                    dwgDIBA.Add(diba);
                }

                if (dwgDIBA.Count == dbDIBA.Count)
                {
                    for (int row = 0; row < dwgDIBA.Count; row++)
                    {
                        DistintaBase dwg = dwgDIBA[row];
                        DistintaBase dbD = dbDIBA[row];

                        bool eq = checkDIBA(dwg, dbD);
                    }

                }

                Console.WriteLine("a");
            }
        }
        private static bool checkDIBA(DistintaBase d1, DistintaBase d2)
        {
            if (d1.CODICE != d2.CODICE) return false;
            if (d1.DESCRIZIONE != d2.DESCRIZIONE) return false;
            if (d1.OC != d2.OC) return false;
            if (d1.SOURCE != d2.SOURCE) return false;
            if (d1.UM != d2.UM) return false;
            if (d1.TIPOLOGIA != d2.TIPOLOGIA) 
            {
                d1.TIPOLOGIA = d1.TIPOLOGIA.Replace("200000", "");
                d2.TIPOLOGIA = d1.TIPOLOGIA.Replace("200000", "");

                if (d1.TIPOLOGIA != d2.TIPOLOGIA) return false;
            } 
            if (d1.QTA != d2.QTA)
            {
                double qt1 = double.Parse(d1.QTA) * 10000;
                double qt2 = double.Parse(d2.QTA) * 10000;

                double delta = Math.Abs(qt1 - qt2);

                if (delta > 200)
                {
                    return false;
                }
            };


            return true;
        }
    }
}