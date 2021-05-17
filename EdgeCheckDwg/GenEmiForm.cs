using Autodesk.AutoCAD.DatabaseServices;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace EdgeAutocadPlugins
{
    public partial class GenEmiForm : Form
    {
        public static string codiceCommessa = "200000";

        public GenEmiForm()
        {
            InitializeComponent();
        }

        public static string nomeTemplate = "TemplateListaFocchi.xlsm";
        public static string cartellaEmissione = @"";

        public static String CleanHeaderName(String szIn)
        {
            String result = "";

            // elimina righe oltre la prima (commenti, u.m.)
            int r = szIn.IndexOf("\\P");
            if (r > 0) szIn = szIn.Substring(0, r);

            foreach (Char c in szIn)
            {
                if (Char.IsLetterOrDigit(c)) result += c;
                else if ((c == '(') || (c == '[') || (c == '{')) break;
            }

            return result;
        }

        public static String StripFormatString(String szIn)
        {
            String result = szIn;
            int p1, p2, f1, f2;

            do
            {
                // cerca delimitatori testo formattato
                p1 = result.IndexOf('{');
                p2 = result.IndexOf('}');
                if ((p1 == -1) || (p2 == -1)) break;

                if (p2 > p1)
                {
                    String fld = result.Substring(p1 + 1, p2 - p1 - 1);

                    // cerca delimitatori formattazione e rimuovi
                    do
                    {
                        f1 = fld.IndexOf('\\');
                        f2 = fld.IndexOf(';');

                        if ((f1 == -1) || (f2 == -1)) break;
                        if (f2 > f1) fld = fld.Remove(f1, f2 - f1 + 1);
                    }
                    while (f2 > f1);

                    // rimuovi testo delimitato da { } e rimpiazza con contenuto depurato
                    result = result.Remove(p1, p2 - p1 + 1);
                    result = result.Insert(p1, fld);
                }
            }
            while (p2 > p1);

            // la coppia '\\' e ';' può presentarsi anche senza parentesi graffe
            // cerca delimitatori formattazione e rimuovi
            do
            {
                f1 = result.IndexOf('\\');
                f2 = result.IndexOf(';');

                if ((f1 == -1) || (f2 == -1)) break;
                if (f2 > f1) result = result.Remove(f1, f2 - f1 + 1);
            }
            while (f2 > f1);

            return result;
        }

        public static System.Data.DataTable risultatiTabella(string nomeFile)
        {
            Database db = new Database(false, true);
            db.ReadDwgFile(nomeFile, FileOpenMode.OpenForReadAndAllShare, false, null);

            ObjectId result = new ObjectId();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("CODICEOP", Type.GetType("System.String"));
            dt.Columns.Add("DISTANZAI", Type.GetType("System.String"));
            dt.Columns.Add("DESCRIZIONEOP", Type.GetType("System.String"));
            dt.Columns.Add("DISTI", Type.GetType("System.String"));
            dt.Columns.Add("nrow", Type.GetType("System.Int32"));

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId idTable = FindTable(db, tr, "OPERAZ");

                if (idTable != ObjectId.Null)
                {
                    Table tbl = (Table)tr.GetObject(idTable, OpenMode.ForRead);

                    int[] Indexes = new int[100];

                    for (int i = 0; i < Indexes.Length; i++) Indexes[i] = -1;

                    // trovata tabella Distinta Base
                    // leggi e trasferisci tutte le righe in tabella di appoggio
                    for (int h = 0; h < tbl.Rows.Count; h++)
                    {
                        String szRowType = tbl.Cells[h, -1].Style.ToUpper();

                        if (szRowType == @"_HEADER")
                        {
                            // analizza titoli colonne per terminare quali colonne
                            // caricare e associale alla datarow corrente
                            for (int j = 0; j < tbl.Columns.Count; j++)
                            {
                                String sz = tbl.Cells[h, j].GetTextString(FormatOption.IgnoreMtextFormat).ToUpper();    // recupera titolo
                                sz = CleanHeaderName(sz);
                                if (sz.Length == 0) continue;

                                // abbina a indice colonna in DataTable
                                if (sz.StartsWith(@"CODICE")) Indexes[j] = 0;
                                else if (sz == @"DIST") Indexes[j] = 1;
                                else if (sz.StartsWith(@"DESCRIZION")) Indexes[j] = 2;
                                else if (sz == @"DISTINV") Indexes[j] = 3;
                            }
                        }
                        else if (szRowType == @"_DATA")
                        {
                            DataRow dr_new = dt.NewRow();
                            dr_new["nrow"] = h;

                            // in base ai titoli colonne determina quale campo caricare
                            for (int j = 0; j < tbl.Columns.Count; j++)
                            {
                                if (Indexes[j] > -1)
                                {
                                    // colonna associata, trasferisci
                                    dr_new[Indexes[j]] = StripFormatString(tbl.Cells[h, j].TextString);
                                }
                            }
                            if (!string.IsNullOrEmpty(dr_new.ItemArray[0].ToString()))
                                dt.Rows.Add(dr_new);
                        }
                    }

                    // salva modifiche a tabella
                    dt.AcceptChanges();
                }

            }

            db.Dispose();

            return dt;
        }

        public static ObjectId FindTable(Database dwg, Transaction tr, String szHeaderKey)
        {
            ObjectId result = ObjectId.Null;

            try
            {
                // separa eventuali chiavi multiple
                String[] vKeys = szHeaderKey.ToUpper().Split('|');

                BlockTable bt = (BlockTable)tr.GetObject(dwg.BlockTableId, OpenMode.ForRead);
                String szFilename = Path.GetFileNameWithoutExtension(dwg.Filename);

                // cerca in tutti i layouts
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    // se è un layout, ma non è lo spazio Modello, aggiungi a lista
                    if (btr.IsLayout && (btr.Name.ToUpper() != BlockTableRecord.ModelSpace.ToUpper()))
                    {
                        result = _FindTable(btr, tr, vKeys);

                        if (result != ObjectId.Null) break;
                    }
                }

                //BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);
                //result = _FindTable(ms, tr, vKeys);

                if (result == ObjectId.Null)
                {
                    // non trovato in paperspace, prova in modelspace
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    result = _FindTable(ms, tr, vKeys);
                }
            }
            catch { }

            return result;
        }

        public static ObjectId _FindTable(BlockTableRecord ms, Transaction tr, String[] vKeys)
        {
            ObjectId result = ObjectId.Null;

            foreach (ObjectId o in ms)
            {
                DBObject dbo = tr.GetObject(o, OpenMode.ForRead);
                if ((dbo is Table) && (!dbo.IsErased))
                {
                    Table tbl = dbo as Table;

                    String szTitle = tbl.Cells[0, 0].TextString.ToUpper();

                    foreach (String s in vKeys)
                    {
                        if (szTitle.Contains(s))
                        {
                            result = o;
                            break;
                        }
                    }
                }
            }

            return result;
        }

        public static string[] cleanModelli(string modelli)
        {
            string[] stringSeparators = new string[] { "\r\n" };

            string[] res = modelli.Split(stringSeparators, StringSplitOptions.None);

            int ind = 0;

            res = res.Where(o => o != "").ToArray();

            return res;
        }

        private void start_Click(object sender, EventArgs e)
        {
            // ! VALIDATORE FOLDER
            if (string.IsNullOrEmpty(cartellaEmissione))
            {
                MessageBox.Show("Selezionare cartella emissione valida", "Errore");
                return;
            }

            string[] list = Directory.GetFiles(cartellaEmissione, "*.dwg", SearchOption.AllDirectories);

            progressBar1.Value = 0;
            progressBar1.Maximum = 3;

            // Audit Dwg
            if (auditCheckBox.Checked)
            {
                //Commands.auditDwg(list);
            }

            // GENERAZIONE EXCEL
            if (arasExcelCheckbox.Checked)
            {

                // ! Prendo le informazioni sul file di template
                FileInfo fiTemplate = new FileInfo(@"C:\Users\edgesuser\source\repos\GeneraEmissione\GeneraEmissione\" + nomeTemplate);
                // ! Genero il file excel
                FileInfo fiNewFile = generaTemplateExcel();


                // ! Suddivido i dwg per classe all'interno delle cartelle
                Dictionary<string, List<string>> dslsClasseFiles = gestiscoFiles(cartellaEmissione);


                // ! Prendo i codici CE nel caso ci fossero per eventuali pezzi distribuiti dentro il database
                List<string> lsCodiciCellula = new List<string>();
                if (dslsClasseFiles.ContainsKey("CE") && dslsClasseFiles["CE"].Count > 0)
                {
                    lsCodiciCellula = codiceWithoutRevisione(dslsClasseFiles["CE"]);
                }


                // ! Inizio a processare i dwg [CELLULE]
                List<ceC> lces = new List<ceC>();
                if (dslsClasseFiles.ContainsKey("CE") && dslsClasseFiles["CE"].Count > 0)
                    lces = Commands.CicloCE(dslsClasseFiles["CE"]);



                // ! Inizio a processare i dwg [OPERAZIONI]
                List<opC> lops = new List<opC>();
                if (dslsClasseFiles.ContainsKey("OP") && dslsClasseFiles["OP"].Count > 0)
                    lops = Commands.CicloOP(dslsClasseFiles["OP"]);




                // ! Tutto ciò che non è CE e PE
                // ! Faccio un vettore con tutti i dwg
                List<string> lsFiles = new List<string>();

                foreach (KeyValuePair<string, List<string>> kvp in dslsClasseFiles)
                {
                    List<string> lFiles = kvp.Value;

                    if (lFiles.Count > 0) lsFiles.AddRange(lFiles);
                }

                // ! Inizio a processare i dwg [PE-AS-GU.....]
                List<peC> lpes = new List<peC>();
                if (lsFiles.Count > 0)
                    lpes = Commands.CicloPE(lsFiles);



                // ! Excel
                if (fiTemplate != null || fiNewFile != null)
                {
                    using (ExcelPackage xlPackage = new ExcelPackage(fiNewFile, fiTemplate))
                    {
                        // ! Anagrafica
                        writeAnagraficaExcel(lpes, lces, xlPackage);


                        // ! Operazioni
                        writeOperazioniExcel(lops, xlPackage);


                        // ! Distinta Base
                        writeDistintaBaseExcel(lpes, lces, xlPackage);



                        // ! Impegnati
                        writeImpegnatiExcel(lsCodiciCellula, xlPackage);



                        // ! Salvataggio Excel
                        xlPackage.Save();
                    }


                    //FileStream fs = File.Create(cartellaEmissione+"\\EmissioneAras.xlsm");
                    //fs.Close();

                    //FileInfo newFile = new FileInfo(cartellaEmissione + "\\EmissioneAras.xlsm");

                    //FileInfo template = new FileInfo(@"C:\Users\edgesuser\source\repos\GeneraEmissione\GeneraEmissione\" + nomeTemplate);

                    //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    //// ANAGRAFICA && DIBA
                    //// ? leggere info da cartiglio

                    //    if (radioButton3.Checked)
                    //    {
                    //        if (Directory.Exists(cartellaEmissione + "\\CE") && Directory.Exists(cartellaEmissione + "\\PE"))
                    //        {
                    //            string[] listCE = Directory.GetFiles(cartellaEmissione + "\\CE", "*.dwg");
                    //            string[] listPE = Directory.GetFiles(cartellaEmissione + "\\PE", "*.dwg");

                    //            List<string> codiciCE = new List<string>();

                    //            foreach (string nf in listCE)
                    //            {
                    //                string clnf = Path.GetFileName(nf).Replace(codiceCommessa, "").Split('_')[0] + "-001";
                    //                codiciCE.Add(clnf);
                    //            }

                    //            List<ceC> ces = Commands.CicloCE(listCE);
                    //            List<peC> pes = Commands.CicloPE(listPE);

                    //            using (ExcelPackage xlPackage = new ExcelPackage(newFile, template))
                    //            {

                    //                int index = 2;
                    //                string nameWorkBook = "Anagrafica";

                    //                ExcelWorksheet aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                    //                foreach (ceC cellula in ces)
                    //                {
                    //                    foreach (DataRow r in cellula.POSIZIONI.Rows)
                    //                    {
                    //                        try
                    //                        {
                    //                            aworksheet.Cells[index, 1].Value = cellula.COMMESSA.ToString() + cellula.NOME.ToString() + "-" + r[0].ToString();
                    //                            aworksheet.Cells[index, 2].Value = cellula.COMMESSA.ToString();
                    //                            aworksheet.Cells[index, 3].Value = cellula.CLASSE.ToString();
                    //                            aworksheet.Cells[index, 4].Value = cellula.SOTTOCLASSE.ToString();
                    //                            aworksheet.Cells[index, 5].Value = cellula.NOME.ToString().Substring(2, 4);
                    //                            aworksheet.Cells[index, 6].Value = r[0].ToString();
                    //                            aworksheet.Cells[index, 7].Value = cellula.DESCRIZIONE.ToString();
                    //                            aworksheet.Cells[index, 11].Value = (cellula.TYPE == "CANTIERE" ? "C" : "O"); // Destinazione
                    //                            aworksheet.Cells[index, 12].Value = ""; // Materiale
                    //                            aworksheet.Cells[index, 13].Value = ""; // Finitura
                    //                            aworksheet.Cells[index, 15].Value = float.Parse(r[1].ToString());
                    //                            aworksheet.Cells[index, 16].Value = float.Parse(r[2].ToString());
                    //                            aworksheet.Cells[index, 18].Value = cellula.EMESSO.ToString();
                    //                            aworksheet.Cells[index, 19].Value = cellula.NOMEEMI; // Descrizione rev
                    //                            aworksheet.Cells[index, 21].Value = cellula.NOMEDWG.ToString();
                    //                            aworksheet.Cells[index, 22].Value = cellula.NOMEPDF.ToString();
                    //                            index += 1;
                    //                        }
                    //                        catch { }
                    //                    }
                    //                }

                    //                foreach (peC pe in pes)
                    //                {
                    //                    foreach (DataRow r in pe.POSIZIONI.Rows)
                    //                    {
                    //                        try
                    //                        {
                    //                            aworksheet.Cells[index, 1].Value = pe.COMMESSA.ToString() + pe.NOME.ToString() + "-" + r[0].ToString();
                    //                            aworksheet.Cells[index, 2].Value = pe.COMMESSA.ToString();
                    //                            aworksheet.Cells[index, 3].Value = pe.CLASSE.ToString();
                    //                            aworksheet.Cells[index, 4].Value = pe.SOTTOCLASSE.ToString();
                    //                            aworksheet.Cells[index, 5].Value = pe.NOME.ToString().Substring(2, 4);
                    //                            aworksheet.Cells[index, 6].Value = r[0].ToString();
                    //                            aworksheet.Cells[index, 7].Value = pe.DESCRIZIONE.ToString();
                    //                            aworksheet.Cells[index, 11].Value = (pe.TYPE == "CANTIERE" ? "C" : "O"); // Destinazione


                    //                            // Devo cercare cosa c'è tra parentesi
                    //                            aworksheet.Cells[index, 12].Value = GenericFunction.findParentesi(pe.MATERIALE); // Materiale
                    //                            aworksheet.Cells[index, 13].Value = GenericFunction.findParentesi(pe.TRATTAMENTO); // Finitura

                    //                            try
                    //                            {
                    //                                aworksheet.Cells[index, 15].Value = float.Parse(r[1].ToString());
                    //                            }
                    //                            catch { }
                    //                            aworksheet.Cells[index, 16].Value = r[2].ToString();
                    //                            aworksheet.Cells[index, 18].Value = pe.EMESSO.ToString();
                    //                            aworksheet.Cells[index, 19].Value = pe.NOMEEMI; // Descrizione rev
                    //                            aworksheet.Cells[index, 21].Value = pe.NOMEDWG.ToString();
                    //                            aworksheet.Cells[index, 22].Value = pe.NOMEPDF.ToString();
                    //                            index += 1;
                    //                        }
                    //                        catch { }
                    //                    }
                    //                }


                    //                index = 2;
                    //                nameWorkBook = "Distinta Base";

                    //                aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                    //                foreach (ceC cellula in ces)
                    //                {
                    //                    // RICORDARSI DI ORDINARE

                    //                    //DataRowCollection drc = cellula.DIBA.Rows;
                    //                    int indexConteggio = 1;
                    //                    foreach (DataRow r in cellula.DIBA.Rows)
                    //                    {
                    //                        try
                    //                        {
                    //                            aworksheet.Cells[index, 1].Value = indexConteggio;
                    //                            aworksheet.Cells[index, 2].Value = r[0].ToString();
                    //                            aworksheet.Cells[index, 3].Value = r[1].ToString();
                    //                            aworksheet.Cells[index, 4].Value = r[5].ToString();
                    //                            aworksheet.Cells[index, 5].Value = float.Parse(r[2].ToString());
                    //                            aworksheet.Cells[index, 6].Value = r[4].ToString();
                    //                            index += 1;

                    //                            indexConteggio += 1;
                    //                        }
                    //                        catch { }
                    //                    }

                    //                    index += 1;
                    //                }

                    //                foreach (peC pe in pes)
                    //                {
                    //                    int indexConteggio = 1;
                    //                    if (pe.DIBA == null) continue;
                    //                    foreach (DataRow r in pe.DIBA.Rows)
                    //                    {
                    //                        try
                    //                        {
                    //                            aworksheet.Cells[index, 1].Value = indexConteggio;
                    //                            aworksheet.Cells[index, 2].Value = r[0].ToString();
                    //                            aworksheet.Cells[index, 3].Value = r[1].ToString();
                    //                            aworksheet.Cells[index, 4].Value = r[5].ToString();
                    //                            aworksheet.Cells[index, 5].Value = float.Parse(r[2].ToString());
                    //                            aworksheet.Cells[index, 6].Value = r[4].ToString();
                    //                            index += 1;

                    //                            indexConteggio += 1;
                    //                        }
                    //                        catch { }
                    //                    }

                    //                    index += 1;
                    //                }



                    //                // OP
                    //                // ? leggere info da cartiglio

                    //                if (Directory.Exists(cartellaEmissione + "\\OP"))
                    //                {
                    //                    string[] listOP = Directory.GetFiles(cartellaEmissione + "\\OP", "*.dwg");
                    //                    // devo instanziare autocad e processare

                    //                    List<opC> ops = Commands.CicloOP(listOP);
                    //                    index = 2;
                    //                    nameWorkBook = "OP";

                    //                    aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                    //                    foreach (opC op in ops)
                    //                    {
                    //                        try
                    //                        {
                    //                            aworksheet.Cells[index, 1].Value = op.CODICE_PROFILO + "-" + op.N_OP;
                    //                            aworksheet.Cells[index, 2].Value = op.CODICE_PROFILO;
                    //                            aworksheet.Cells[index, 3].Value = op.N_OP.Replace("OP", "");
                    //                            aworksheet.Cells[index, 4].Value = op.DESCRIZIONE;
                    //                            aworksheet.Cells[index, 5].Value = "di Commessa";
                    //                            aworksheet.Cells[index, 6].Value = codiceCommessa;
                    //                            aworksheet.Cells[index, 7].Value = op.NOMEFILEDWG;
                    //                            aworksheet.Cells[index, 8].Value = op.NOMEFILEPDF;

                    //                            index += 1;
                    //                        }
                    //                        catch{}
                    //                    }
                    //                }


                    //                // IMPEGNATI
                    //                // ? leggere info da database

                    //                try
                    //                {
                    //                    Distribuzione d = new Distribuzione();

                    //                    Dictionary<string, dynamic> _results = new Dictionary<string, dynamic>();

                    //                    if (listaRadio.Checked)
                    //                        _results = d.customQueryLotID(codiciCE);
                    //                    else if (pianiRadio.Checked)
                    //                        _results = d.customQueryFasciePiani(codiciCE);

                    //                    var l = _results.OrderBy(key => key.Key);
                    //                    Dictionary<string, dynamic> results = l.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);

                    //                    index = 2;
                    //                    nameWorkBook = "Impegnati";

                    //                    aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                    //                    foreach (KeyValuePair<string, dynamic> data in results)
                    //                    {
                    //                        string lotto = data.Key;

                    //                        int indexConteggio = 1;

                    //                        foreach (dynamic count in data.Value)
                    //                        {
                    //                            string nomeLotto = "";

                    //                            string tipologia = lotto.Substring(1, 1);

                    //                            if (tipologia == "B")
                    //                                nomeLotto = "Brise Soleil Units";
                    //                            if (tipologia == "K")
                    //                                nomeLotto = "Corner Units";
                    //                            if (tipologia == "C")
                    //                                nomeLotto = "Casing Units";
                    //                            if (tipologia == "A")
                    //                                nomeLotto = "Standard Units";
                    //                            if (tipologia == "H")
                    //                                nomeLotto = "Hoist Infill";
                    //                            if (tipologia == "R")
                    //                                nomeLotto = "Roof";
                    //                            if (tipologia == "P")
                    //                                nomeLotto = "Parapets";

                    //                            aworksheet.Cells[index, 1].Value = indexConteggio;
                    //                            aworksheet.Cells[index, 2].Value = codiceCommessa + "LIS" + lotto.ToString();
                    //                            aworksheet.Cells[index, 4].Value = nomeLotto;
                    //                            aworksheet.Cells[index, 5].Value = codiceCommessa + count.Value.ToString();
                    //                            aworksheet.Cells[index, 7].Value = count.Count;
                    //                            aworksheet.Cells[index, 8].Value = (rO.Checked ? "O" : "C");

                    //                            indexConteggio += 1;
                    //                            index += 1;
                    //                        }

                    //                        index += 1;
                    //                    }
                    //                }
                    //                catch
                    //                {

                    //                }

                    //                xlPackage.Save();
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        // ! Nel caso nel cartiglio non da programmone [Generico]
                    //        // ! Cerca in tutte le sotto cartelle
                    //        string[] listDWG = Directory.GetFiles(cartellaEmissione, "*.dwg", SearchOption.AllDirectories);
                    //        List<peC> pes = Commands.CicloCartigioGenerico((string[])listDWG.Where(a => !a.Contains("OP")).ToArray());
                    //        List<opC> ops = Commands.CicloOP((string[])listDWG.Where(a => a.Contains("OP")).ToArray());

                    //        using (ExcelPackage xlPackage = new ExcelPackage(newFile, template))
                    //        {

                    //            int index = 2;
                    //            string nameWorkBook = "Anagrafica";

                    //            ExcelWorksheet aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                    //            foreach (peC pe in pes)
                    //            {
                    //                foreach (DataRow r in pe.POSIZIONI.Rows)
                    //                {
                    //                    try
                    //                    {
                    //                        aworksheet.Cells[index, 1].Value = pe.COMMESSA.ToString() + pe.NOME.ToString() + "-" + r[0].ToString();
                    //                        aworksheet.Cells[index, 2].Value = pe.COMMESSA.ToString();
                    //                        aworksheet.Cells[index, 3].Value = pe.CLASSE.ToString();
                    //                        aworksheet.Cells[index, 4].Value = pe.SOTTOCLASSE.ToString();
                    //                        aworksheet.Cells[index, 5].Value = pe.NOME.ToString().Substring(2, 4);
                    //                        aworksheet.Cells[index, 6].Value = r[0].ToString();
                    //                        aworksheet.Cells[index, 7].Value = pe.DESCRIZIONE.ToString();
                    //                        aworksheet.Cells[index, 11].Value = (pe.TYPE == "CANTIERE" ? "C" : "O"); // Destinazione


                    //                        // Devo cercare cosa c'è tra parentesi
                    //                        aworksheet.Cells[index, 12].Value = GenericFunction.findParentesi(pe.MATERIALE); // Materiale
                    //                        aworksheet.Cells[index, 13].Value = GenericFunction.findParentesi(pe.TRATTAMENTO); // Finitura

                    //                        try
                    //                        {
                    //                            aworksheet.Cells[index, 15].Value = float.Parse(r[1].ToString());
                    //                        }
                    //                        catch { }
                    //                        aworksheet.Cells[index, 16].Value = r[2].ToString();
                    //                        aworksheet.Cells[index, 18].Value = pe.EMESSO.ToString();
                    //                        aworksheet.Cells[index, 19].Value = pe.NOMEEMI; // Descrizione rev
                    //                        aworksheet.Cells[index, 21].Value = pe.NOMEDWG.ToString();
                    //                        aworksheet.Cells[index, 22].Value = pe.NOMEPDF.ToString();
                    //                        index += 1;
                    //                    }
                    //                    catch { }
                    //                }
                    //            }

                    //            index = 2;
                    //            nameWorkBook = "OP";

                    //            aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                    //            foreach (opC op in ops)
                    //            {
                    //                try
                    //                {
                    //                    aworksheet.Cells[index, 1].Value = op.CODICE_PROFILO + "-" + op.N_OP;
                    //                    aworksheet.Cells[index, 2].Value = op.CODICE_PROFILO;
                    //                    aworksheet.Cells[index, 3].Value = op.N_OP.Replace("OP", "");
                    //                    aworksheet.Cells[index, 4].Value = op.DESCRIZIONE;
                    //                    aworksheet.Cells[index, 5].Value = "di Commessa";
                    //                    aworksheet.Cells[index, 6].Value = codiceCommessa;
                    //                    aworksheet.Cells[index, 7].Value = op.NOMEFILEDWG;
                    //                    aworksheet.Cells[index, 8].Value = op.NOMEFILEPDF;

                    //                    index += 1;
                    //                }
                    //                catch{}
                    //            }
                    //            xlPackage.Save();
                    //        }
                    //    }
                    //}
                }

                progressBar1.Value = 1;

                // GENERAZIONE PDF
                if (pdfCheckbox.Checked)
                {
                    Commands.generoPDF(list);
                }

                progressBar1.Value = 2;

                // PUBBLICAZIONE LOGICA DWG
                if (pubblicaCheckbox.Checked)
                {
                }

                progressBar1.Value = 3;
            }
        }

        private void writeImpegnatiExcel(List<string> codiciCE, ExcelPackage xlPackage)
        {
            try
            {
                Distribuzione d = new Distribuzione();

                Dictionary<string, dynamic> _results = new Dictionary<string, dynamic>();

                if (listaRadio.Checked)
                    _results = d.customQueryLotID(codiciCE);
                else if (pianiRadio.Checked)
                    _results = d.customQueryFasciePiani(codiciCE);

                var l = _results.OrderBy(key => key.Key);
                Dictionary<string, dynamic> results = l.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);

                int index = 2;
                string nameWorkBook = "Impegnati";

                ExcelWorksheet aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

                foreach (KeyValuePair<string, dynamic> data in results)
                {
                    string lotto = data.Key;

                    int indexConteggio = 1;

                    foreach (dynamic count in data.Value)
                    {
                        string nomeLotto = "";

                        string tipologia = lotto.Substring(1, 1);

                        if (tipologia == "B")
                            nomeLotto = "Brise Soleil Units";
                        if (tipologia == "K")
                            nomeLotto = "Corner Units";
                        if (tipologia == "C")
                            nomeLotto = "Casing Units";
                        if (tipologia == "A")
                            nomeLotto = "Standard Units";
                        if (tipologia == "H")
                            nomeLotto = "Hoist Infill";
                        if (tipologia == "R")
                            nomeLotto = "Roof";
                        if (tipologia == "P")
                            nomeLotto = "Parapets";

                        aworksheet.Cells[index, 1].Value = indexConteggio;
                        aworksheet.Cells[index, 2].Value = codiceCommessa + "LIS" + lotto.ToString();
                        aworksheet.Cells[index, 4].Value = nomeLotto;
                        aworksheet.Cells[index, 5].Value = codiceCommessa + count.Value.ToString();
                        aworksheet.Cells[index, 7].Value = count.Count;
                        aworksheet.Cells[index, 8].Value = (rO.Checked ? "O" : "C");

                        indexConteggio += 1;
                        index += 1;
                    }

                    index += 1;
                }
            }
            catch
            {

            }
        }

        private void writeDistintaBaseExcel(List<peC> lpes, List<ceC> lces, ExcelPackage xlPackage)
        {
            int index = 2;
            string nameWorkBook = "Distinta Base";

            ExcelWorksheet aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

            foreach (ceC cellula in lces)
            {
                // RICORDARSI DI ORDINARE

                //DataRowCollection drc = cellula.DIBA.Rows;
                int indexConteggio = 1;
                foreach (DataRow r in cellula.DIBA.Rows)
                {
                    try
                    {
                        aworksheet.Cells[index, 1].Value = indexConteggio;
                        aworksheet.Cells[index, 2].Value = r[0].ToString();
                        aworksheet.Cells[index, 3].Value = r[1].ToString();
                        aworksheet.Cells[index, 4].Value = r[5].ToString();
                        aworksheet.Cells[index, 5].Value = float.Parse(r[2].ToString());
                        aworksheet.Cells[index, 6].Value = r[4].ToString();
                        index += 1;

                        indexConteggio += 1;
                    }
                    catch { }
                }

                index += 1;
            }

            foreach (peC pe in lpes)
            {
                int indexConteggio = 1;
                if (pe.DIBA == null) continue;
                foreach (DataRow r in pe.DIBA.Rows)
                {
                    try
                    {
                        aworksheet.Cells[index, 1].Value = indexConteggio;
                        aworksheet.Cells[index, 2].Value = r[0].ToString();
                        aworksheet.Cells[index, 3].Value = r[1].ToString();
                        aworksheet.Cells[index, 4].Value = r[5].ToString();
                        aworksheet.Cells[index, 5].Value = float.Parse(r[2].ToString());
                        aworksheet.Cells[index, 6].Value = r[4].ToString();
                        index += 1;

                        indexConteggio += 1;
                    }
                    catch { }
                }

                index += 1;
            }
        }

        private void writeOperazioniExcel(List<opC> lops, ExcelPackage xlPackage)
        {
            int index = 2;
            string nameWorkBook = "OP";

            ExcelWorksheet aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];

            foreach (opC op in lops)
            {
                try
                {
                    aworksheet.Cells[index, 1].Value = op.CODICE_PROFILO + "-" + op.N_OP;
                    aworksheet.Cells[index, 2].Value = op.CODICE_PROFILO;
                    aworksheet.Cells[index, 3].Value = op.N_OP.Replace("OP", "");
                    aworksheet.Cells[index, 4].Value = op.DESCRIZIONE;
                    aworksheet.Cells[index, 5].Value = "di Commessa";
                    aworksheet.Cells[index, 6].Value = codiceCommessa;
                    aworksheet.Cells[index, 7].Value = op.NOMEFILEDWG;
                    aworksheet.Cells[index, 8].Value = op.NOMEFILEPDF;

                    index += 1;
                }
                catch { }
            }
        }

        private void writeAnagraficaExcel(List<peC> lpes, List<ceC> lces, ExcelPackage xlPackage)
        {
            int index = 2;
            string nameWorkBook = "Anagrafica";

            ExcelWorksheet aworksheet = xlPackage.Workbook.Worksheets[nameWorkBook];


            foreach (ceC cellula in lces.OrderBy(f => f.NOME).ToList())
            {
                foreach (DataRow r in cellula.POSIZIONI.Rows)
                {
                    try
                    {
                        aworksheet.Cells[index, 1].Value = cellula.COMMESSA.ToString() + cellula.NOME.ToString() + "-" + r[0].ToString();
                        aworksheet.Cells[index, 2].Value = cellula.COMMESSA.ToString();
                        aworksheet.Cells[index, 3].Value = cellula.CLASSE.ToString();
                        aworksheet.Cells[index, 4].Value = cellula.SOTTOCLASSE.ToString();
                        aworksheet.Cells[index, 5].Value = cellula.NOME.ToString().Substring(2, 4);
                        aworksheet.Cells[index, 6].Value = r[0].ToString();
                        aworksheet.Cells[index, 7].Value = cellula.DESCRIZIONE.ToString();
                        aworksheet.Cells[index, 11].Value = (cellula.TYPE == "CANTIERE" ? "C" : "O"); // Destinazione
                        aworksheet.Cells[index, 12].Value = ""; // Materiale
                        aworksheet.Cells[index, 13].Value = ""; // Finitura
                        aworksheet.Cells[index, 15].Value = float.Parse(r[1].ToString());
                        aworksheet.Cells[index, 16].Value = float.Parse(r[2].ToString());
                        aworksheet.Cells[index, 18].Value = cellula.EMESSO.ToString();
                        aworksheet.Cells[index, 19].Value = cellula.NOMEEMI; // Descrizione rev
                        aworksheet.Cells[index, 21].Value = cellula.NOMEDWG.ToString();
                        aworksheet.Cells[index, 22].Value = cellula.NOMEPDF.ToString();
                        index += 1;
                    }
                    catch { }
                }
            }

            foreach (peC pe in lpes.OrderBy(f => f.NOME).ToList())
            {
                foreach (DataRow r in pe.POSIZIONI.Rows)
                {
                    try
                    {
                        aworksheet.Cells[index, 1].Value = pe.COMMESSA.ToString() + pe.NOME.ToString() + "-" + r[0].ToString();
                        aworksheet.Cells[index, 2].Value = pe.COMMESSA.ToString();
                        aworksheet.Cells[index, 3].Value = pe.CLASSE.ToString();
                        aworksheet.Cells[index, 4].Value = pe.SOTTOCLASSE.ToString();
                        aworksheet.Cells[index, 5].Value = pe.NOME.ToString().Substring(2, 4);
                        aworksheet.Cells[index, 6].Value = r[0].ToString();
                        aworksheet.Cells[index, 7].Value = pe.DESCRIZIONE.ToString();
                        aworksheet.Cells[index, 11].Value = (pe.TYPE == "CANTIERE" ? "C" : "O"); // Destinazione


                        // Devo cercare cosa c'è tra parentesi
                        aworksheet.Cells[index, 12].Value = GenericFunction.findParentesi(pe.MATERIALE); // Materiale
                        aworksheet.Cells[index, 13].Value = GenericFunction.findParentesi(pe.TRATTAMENTO); // Finitura

                        try
                        {
                            aworksheet.Cells[index, 15].Value = float.Parse(r[1].ToString());
                        }
                        catch { }
                        aworksheet.Cells[index, 16].Value = r[2].ToString();
                        aworksheet.Cells[index, 18].Value = pe.EMESSO.ToString();
                        aworksheet.Cells[index, 19].Value = pe.NOMEEMI; // Descrizione rev
                        aworksheet.Cells[index, 21].Value = pe.NOMEDWG.ToString();
                        aworksheet.Cells[index, 22].Value = pe.NOMEPDF.ToString();
                        index += 1;
                    }
                    catch { }
                }
            }
        }

        private List<string> codiceWithoutRevisione(List<string> list)
        {
            List<string> res = new List<string>();

            foreach (string s in list)
            {

                string sNomeFile = Path.GetFileNameWithoutExtension(s);

                res.Add(sNomeFile.Replace(codiceCommessa, "").Split('_')[0] + "-001");
            }

            return res.Distinct().OrderBy(a => a).ToList();
        }

        private Dictionary<string, List<string>> gestiscoFiles(string sPath)
        {
            // ! Le prime due lettere indicano la classe dell'elemento presente nel dwg 
            // ! Considerare il fatto che ci possa essere il nome della commessa in testa al file
            // ! Il tutto deve essere riconosciuto con delle regex

            Dictionary<string, List<string>> results = new Dictionary<string, List<string>>();

            // ! Prendo tutti i dwg dalla cartella selezionata
            string[] asFiles = Directory.GetFiles(sPath, "*.dwg", SearchOption.AllDirectories);

            foreach (string sNomeFileCompleto in asFiles)
            {
                string sNomeFile = Path.GetFileName(sNomeFileCompleto);

                if (sNomeFile.Length < 2) continue;

                sNomeFile = sNomeFile.Replace(codiceCommessa, "");

                string sClass = sNomeFile.Substring(0, 2);


                // ! Utilizzo le Regex di confronto
                string sRegexOP = @"\d{7}-OP\d{4}_\d{1}.dwg";
                string sRegexStandard = @"[a-zA-Z]{2}\d{4}_\d{1}.dwg";

                RegexOptions options = RegexOptions.Multiline;

                if(Regex.Matches(sNomeFile, sRegexOP, options).Count > 0)
                {
                    // ! Operazione [0102479-OP1000_0.dwg]
                    sClass = "OP";
                }
                else if(Regex.Matches(sNomeFile, sRegexStandard, options).Count > 0)
                {
                    // ! Standard [GU1000_0.dwg]
                    sClass = sClass;
                }



                if (!results.ContainsKey(sClass)) results.Add(sClass, new List<string>());

                results[sClass].Add(sNomeFileCompleto);
            }

            return results;
        }

        private void chooseFolder_Click(object sender, EventArgs e)
        {
            cartellaEmissione = GenericFunction.choosePath();

            labelChooseFolder.Text = cartellaEmissione;
        }

        private void arasExcelCheckbox_CheckedChanged(object sender, EventArgs e)
        {
        }
        private FileInfo generaTemplateExcel()
        {
            // ! Genero file
            FileStream fs = File.Create(cartellaEmissione + "\\EmissioneAras.xlsm");
            fs.Close();

            FileInfo newFile = new FileInfo(cartellaEmissione + "\\EmissioneAras.xlsm");

            // ! Setto licenza free
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            return newFile;
        }
    }
}