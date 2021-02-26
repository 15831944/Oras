using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;

namespace EdgeCheckDwg
{
    public class Commands
    {     
        [CommandMethod("checkPE")]
        public void checkPE()
        {
            string folderInviati = @"X:\Commesse\Focchi\200000 40L\02.Shop_Drawings\02.Inviati\PE";
            string folderRecenti = @"X:\Commesse\Focchi\200000 40L\97 Auto.Cicli\PE";

            Dictionary<string, Dictionary<DataTable, DataTable>> data = new Dictionary<string, Dictionary<DataTable, DataTable>>();

            string[] filePathsRecenti = Directory.GetFiles(folderRecenti, "*.dwg");
            string[] filePathsInviati = Directory.GetFiles(folderInviati, "*.dwg", SearchOption.AllDirectories);

            List<string> filtro = new List<string>()
            {
            };

            filtro = filtro.Distinct().ToList();

            ConfrontoDwg pb = new ConfrontoDwg();
            pb.Show();

            pb.progressBar1.Minimum = 0;

            if (filtro.Count > 0)
                pb.progressBar1.Maximum = filtro.Count - 1;
            else
                pb.progressBar1.Maximum = filePathsRecenti.Length - 1;

            int c = 0;
            foreach (string f in filePathsRecenti)
            {
                string nf = Path.GetFileName(f);
                nf = nf.Replace("200000", "");
                nf = nf.Split('_')[0];

                if (filtro.Count > 0)
                    if (!filtro.Contains(nf))
                        continue;

                //pb.progressBar1.Value = c-1;
                pb.label2.Text = f;

                DataTable dt1 = risultatiTabella(f);

                string nameFile = Path.GetFileName(f);

                string pathInv = filePathsInviati.Where(a => a.Contains(nameFile)).FirstOrDefault();

                if (pathInv != null)
                {
                    DataTable dt2 = risultatiTabella(pathInv);

                    if (!AreTablesTheSame(dt1, dt2))
                    {
                        // Tabelle non uguali
                        Dictionary<DataTable, DataTable> dtConf = new Dictionary<DataTable, DataTable>();
                        dtConf.Add(dt1, dt2);
                        data.Add(f, dtConf);
                    }
                }
                c += 1;
            }

            pb.Dispose();

            List<string> disegniCambiati = data.Keys.ToList();
        }


        public static GenEmiForm formEmi;
        [CommandMethod("generaEmi")]
        public void generaEmi()
        {
            formEmi = new GenEmiForm();
            formEmi.Show();
        }

        public static DataTable risultatiTabellaCE(string nomeFile)
        {
            Database db = new Database(false, true);
            db.ReadDwgFile(nomeFile, FileOpenMode.OpenForReadAndAllShare, false, null);

            DocumentCollection docMgr = Application.DocumentManager;

            ObjectId result = ObjectId.Null;
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("TIPOLOGIA", Type.GetType("System.String"));
            dt.Columns.Add("CODICE", Type.GetType("System.String"));
            dt.Columns.Add("QUANT", Type.GetType("System.String"));
            dt.Columns.Add("UM", Type.GetType("System.String"));
            dt.Columns.Add("nrow", Type.GetType("System.Int32"));

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId idTable = FindTable(db, tr, "DIST");

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

            //oDoc.CloseAndDiscard();

            db.Dispose();

            return dt;
        }
        public static DataTable risultatiTabella(string nomeFile)
        {
            //Document oDoc = Application.DocumentManager.Open(nomeFile);

            //Database db = oDoc.Database;

            Database db = new Database(false, true);
            db.ReadDwgFile(nomeFile, FileOpenMode.OpenForReadAndAllShare, false, null);

            DocumentCollection docMgr = Application.DocumentManager;

            //docMgr.MdiActiveDocument = oDoc;

            ObjectId result = ObjectId.Null;
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

            //oDoc.CloseAndDiscard();

            db.Dispose();

            return dt;
        }
        public static String ToString(Object o)
        {
            String result = "";

            if (o == null) return result;
            if (o is DBNull) return result;

            try
            {
                result = Convert.ToString(o);
            }
            catch
            {

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
        public static bool AreTablesTheSame(DataTable tbl1, DataTable tbl2)
        {
            if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
                return false;


            for (int i = 0; i < tbl1.Rows.Count; i++)
            {
                for (int c = 0; c < tbl1.Columns.Count; c++)
                {
                    var a = tbl1.Rows[i][c];
                    var b = tbl2.Rows[i][c];
                    if (!Equals(tbl1.Rows[i][c], tbl2.Rows[i][c]))
                        return false;
                }
            }
            return true;
        }

        internal static void generoPDF(string[] listaFiles)
        {
            Application.MainWindow.WindowState = Window.State.Minimized;
            Application.MainWindow.Visible = false;

            foreach (string nomefile in listaFiles)
            {
                DocumentCollection acDocMgr = Application.DocumentManager;
                Document acDoc = acDocMgr.Add(nomefile);
                DocumentLock acLckDoc = acDoc.LockDocument();
                
                Application.MainWindow.Visible = false;

                ChangeNonPrintableLayerVisibility(acDoc, false);

                Database db = acDoc.Database;

                _DoPlot(db.OriginalFileName, db, Path.GetDirectoryName(nomefile), "DWG To PDF.pc3", "A3-A4 PLOT_OLIVETTI_Automanzione_Temp.ctb", "ISO full bleed A3 (420.00 x 297.00 MM)", "Scale to Fit", true, false);

                acDoc.CloseAndDiscard();
            }

            Application.MainWindow.Visible = true;
            Application.MainWindow.WindowState = Window.State.Maximized;
        }

        internal static List<opC> CicloOP(string[] listaFiles)
        {
            List<opC> results = new List<opC>();

            foreach (string nomefile in listaFiles)
            {
                Database db = new Database(false, true);

                db.ReadDwgFile(nomefile, FileOpenMode.OpenForReadAndAllShare, false, null);
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //ObjectId idTable = FindTable(db, tr, "OPERAZ");
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    BlockTableRecord btrMS = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId idBlock in btrMS)
                    {
                        Entity ent = (Entity)tr.GetObject(idBlock, OpenMode.ForRead, false);
                        if (ent.GetType() == typeof(BlockReference))
                        {
                            BlockReference blRef = (BlockReference)ent;

                            if (blRef.Name == "OP")
                            {

                                opC cartiglioOP = new opC();

                                foreach (ObjectId id in blRef.AttributeCollection)
                                {
                                    AttributeReference ar = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);

                                    string tag = ar.Tag;
                                    string value = ar.TextString;

                                    if (tag == "CODICE_PROFILO")
                                        cartiglioOP.CODICE_PROFILO = value;
                                    if (tag == "N°_OP")
                                        cartiglioOP.N_OP = value;
                                    if (tag == "DESCRIZIONE")
                                        cartiglioOP.DESCRIZIONE = value;
                                    if (tag == "DATA_EMISSIONE")
                                        cartiglioOP.DATA_EMISSIONE = value;
                                    if (tag == "DISEGNATORE")
                                        cartiglioOP.DISEGNATORE = value;
                                    if (tag == "VERIFICATA_EMISSIONE_DA")
                                        cartiglioOP.VERIFICATA_EMISSIONE_DA = value;
                                }

                                cartiglioOP.NOMEFILEDWG = Path.GetFileName(nomefile);
                                cartiglioOP.NOMEFILEPDF = Path.GetFileName(nomefile).Split('.')[0].ToString() + ".pdf";

                                results.Add(cartiglioOP);
                            }
                        }
                    }
                }

                db.Dispose();
            }

            return results;
        }

        internal static List<ceC> CicloCE(string[] listaFiles)
        {
            List<ceC> results = new List<ceC>();

            foreach (string nomefile in listaFiles)
            {
                //DocumentCollection acDocMgr = Application.DocumentManager;
                //Document acDoc = acDocMgr.Add(nomefile);
                //DocumentLock acLckDoc = acDoc.LockDocument();

                //ChangeLayerVisibility(acDoc, @"TABELLE", false);
                //ChangeLayerVisibility(acDoc, @"PRELIMINARE", false);

                //ChangeNonPrintableLayerVisibility(acDoc, false);

                //Database db = acDoc.Database;

                Database db = new Database(false, true);

                //HostApplicationServices.WorkingDatabase = db;

                db.ReadDwgFile(nomefile, FileOpenMode.OpenForReadAndAllShare, false, null);

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    BlockTableRecord btrMS = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);

                    // Cartiglio

                    ceC cartiglioCE = new ceC();

                    foreach (ObjectId idBlock in btrMS)
                    {
                        Entity ent = (Entity)tr.GetObject(idBlock, OpenMode.ForRead, false);

                        if (ent.GetType() == typeof(BlockReference))
                        {
                            BlockReference blRef = (BlockReference)ent;

                            if (blRef.Name == "focCART_WS")
                            {

                                foreach (ObjectId id in blRef.AttributeCollection)
                                {
                                    AttributeReference ar = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);

                                    string tag = ar.Tag;
                                    string value = ar.TextString;

                                    if (tag == "TITLE_1") // DESCRIZIONE
                                        cartiglioCE.DESCRIZIONE = value;
                                    if (tag == "REVOPE.0") // EMESSO
                                        cartiglioCE.EMESSO = value;
                                    if (tag == "COMM") // COMMESSA
                                        cartiglioCE.COMMESSA = value;
                                    if (tag == "NR_DIS") // NOME // CLASSE // SOTTOCLASSE
                                        cartiglioCE.NOME = value;
                                    if (tag == "TYPE") // DESTINAZIONE
                                        cartiglioCE.TYPE = value;
                                    if(tag == "REVDESC.0")
                                        cartiglioCE.NOMEEMI = value;

                                }

                                cartiglioCE.CLASSE = cartiglioCE.NOME.Substring(0,2);
                                cartiglioCE.SOTTOCLASSE = cartiglioCE.NOME.Substring(0, 2);
                                cartiglioCE.NOMEDWG = Path.GetFileName(nomefile);
                                cartiglioCE.NOMEPDF = Path.GetFileName(nomefile).Split('.')[0].ToString()+".pdf";

                                results.Add(cartiglioCE);
                            }
                        }
                    }


                    // Posizioni
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("POS", Type.GetType("System.String"));
                    dt.Columns.Add("LM", Type.GetType("System.String"));
                    dt.Columns.Add("HM", Type.GetType("System.String"));
                    dt.Columns.Add("nrow", Type.GetType("System.Int32"));

                    ObjectId idTable = FindTable(db, tr, "POS");
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

                            if (szRowType == @"CELLA POS")
                            {
                                // analizza titoli colonne per terminare quali colonne
                                // caricare e associale alla datarow corrente
                                for (int j = 0; j < tbl.Columns.Count; j++)
                                {
                                    String sz = tbl.Cells[h, j].GetTextString(FormatOption.IgnoreMtextFormat).ToUpper();    // recupera titolo
                                    sz = CleanHeaderName(sz);
                                    if (sz.Length == 0) continue;

                                    // abbina a indice colonna in DataTable
                                    if (sz.StartsWith(@"POS")) Indexes[j] = 0;
                                    else if (sz == @"LM") Indexes[j] = 1;
                                    else if (sz.StartsWith(@"HM")) Indexes[j] = 2;
                                }
                            }
                            else if (szRowType == @"CELLA DATI")
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

                        cartiglioCE.POSIZIONI = dt;
                    }

                    // DIBA

                    System.Data.DataTable dtDIBA = new System.Data.DataTable();
                    dtDIBA.Columns.Add("TIPOLOGIA", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("CODICE", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("QTA", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("UM", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("OC", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("DESCRIZIONE", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("nrow", Type.GetType("System.Int32"));

                    ObjectId idTableDIBA = FindTable(db, tr, "DIST");
                    if (idTable != ObjectId.Null)
                    {
                        Table tbl = (Table)tr.GetObject(idTableDIBA, OpenMode.ForRead);

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
                                    if (sz.StartsWith(@"TIPOLOGIA")) Indexes[j] = 0;
                                    else if (sz == @"CODICE") Indexes[j] = 1;
                                    else if (sz == @"QTA") Indexes[j] = 2;
                                    else if (sz == @"UM") Indexes[j] = 3;
                                    else if (sz == @"OC") Indexes[j] = 4;
                                    else if (sz == @"DESCRIZIONE") Indexes[j] = 5;
                                }
                            }
                            else if (szRowType == @"_DATA")
                            {
                                DataRow dr_new = dtDIBA.NewRow();
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
                                    dtDIBA.Rows.Add(dr_new);
                            }
                        }

                        // salva modifiche a tabella
                        dtDIBA.AcceptChanges();

                        cartiglioCE.DIBA = dtDIBA;
                    }
                }

                db.Dispose();
            }

            return results;
        }

        internal static List<peC> CicloPE(string[] listaFiles)
        { 
            List<peC> results = new List<peC>();

            foreach (string nomefile in listaFiles)
            {
                Database db = new Database(false, true);

                db.ReadDwgFile(nomefile, FileOpenMode.OpenForReadAndAllShare, false, null);
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //ObjectId idTable = FindTable(db, tr, "OPERAZ");
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    BlockTableRecord btrMS = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);

                    // Cartiglio

                    peC cartiglioPE = new peC();

                    foreach (ObjectId idBlock in btrMS)
                    {
                        Entity ent = (Entity)tr.GetObject(idBlock, OpenMode.ForRead, false);

                        if (ent.GetType() == typeof(BlockReference))
                        {
                            BlockReference blRef = (BlockReference)ent;

                            if (blRef.Name == "focCART_WS")
                            {

                                foreach (ObjectId id in blRef.AttributeCollection)
                                {
                                    AttributeReference ar = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);

                                    string tag = ar.Tag;
                                    string value = ar.TextString;

                                    if (tag == "TITLE_1") // DESCRIZIONE
                                        cartiglioPE.DESCRIZIONE = value;
                                    if (tag == "REVOPE.0") // EMESSO
                                        cartiglioPE.EMESSO = value;
                                    if (tag == "COMM") // COMMESSA
                                        cartiglioPE.COMMESSA = value;
                                    if (tag == "NR_DIS") // NOME // CLASSE // SOTTOCLASSE
                                        cartiglioPE.NOME = value;
                                    if (tag == "TYPE") // DESTINAZIONE
                                        cartiglioPE.TYPE = value;
                                    if (tag == "REVDESC.0")
                                        cartiglioPE.NOMEEMI = value;
                                    if (tag == "TRATTAMENTO")
                                        cartiglioPE.TRATTAMENTO = value;
                                    if (tag == "MATERIALE")
                                        cartiglioPE.MATERIALE = value;

                                }

                                cartiglioPE.CLASSE = cartiglioPE.NOME.Substring(0, 2);
                                cartiglioPE.SOTTOCLASSE = cartiglioPE.NOME.Substring(0, 2);
                                cartiglioPE.NOMEDWG = Path.GetFileName(nomefile);
                                cartiglioPE.NOMEPDF = Path.GetFileName(nomefile).Split('.')[0].ToString() + ".pdf";

                                results.Add(cartiglioPE);
                            }
                        }
                    }


                    // Posizioni
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("POS", Type.GetType("System.String"));
                    dt.Columns.Add("LM", Type.GetType("System.String"));
                    dt.Columns.Add("HM", Type.GetType("System.String"));
                    dt.Columns.Add("nrow", Type.GetType("System.Int32"));

                    ObjectId idTable = FindTable(db, tr, "POS");
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

                            if (szRowType == @"CELLA POS")
                            {
                                // analizza titoli colonne per terminare quali colonne
                                // caricare e associale alla datarow corrente
                                for (int j = 0; j < tbl.Columns.Count; j++)
                                {
                                    String sz = tbl.Cells[h, j].GetTextString(FormatOption.IgnoreMtextFormat).ToUpper();    // recupera titolo
                                    sz = CleanHeaderName(sz);
                                    if (sz.Length == 0) continue;

                                    // abbina a indice colonna in DataTable
                                    if (sz.StartsWith(@"POS")) Indexes[j] = 0;
                                    else if (sz == @"LM") Indexes[j] = 1;
                                    else if (sz.StartsWith(@"HM")) Indexes[j] = 2;
                                }
                            }
                            else if (szRowType == @"CELLA DATI")
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

                        cartiglioPE.POSIZIONI = dt;
                    }

                    // DIBA

                    System.Data.DataTable dtDIBA = new System.Data.DataTable();
                    dtDIBA.Columns.Add("TIPOLOGIA", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("CODICE", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("QTA", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("UM", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("OC", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("DESCRIZIONE", Type.GetType("System.String"));
                    dtDIBA.Columns.Add("nrow", Type.GetType("System.Int32"));

                    ObjectId idTableDIBA = FindTable(db, tr, "DIST");
                    if (idTable != ObjectId.Null)
                    {
                        Table tbl = (Table)tr.GetObject(idTableDIBA, OpenMode.ForRead);

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
                                    if (sz.StartsWith(@"TIPOLOGIA")) Indexes[j] = 0;
                                    else if (sz == @"CODICE") Indexes[j] = 1;
                                    else if (sz == @"QTA") Indexes[j] = 2;
                                    else if (sz == @"UM") Indexes[j] = 3;
                                    else if (sz == @"OC") Indexes[j] = 4;
                                    else if (sz == @"DESCRIZIONE") Indexes[j] = 5;
                                }
                            }
                            else if (szRowType == @"_DATA")
                            {
                                DataRow dr_new = dtDIBA.NewRow();
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
                                    dtDIBA.Rows.Add(dr_new);
                            }
                        }

                        // salva modifiche a tabella
                        dtDIBA.AcceptChanges();

                        cartiglioPE.DIBA = dtDIBA;
                    }
                }


                db.Dispose();
            }

            return results;
        }

        
        public static void ChangeNonPrintableLayerVisibility(Document doc, bool bVisible)
        {
            // esegui operazioni sul documento aperto
            Transaction tr = doc.TransactionManager.StartTransaction();

            try
            {
                // recupera tabella layers
                LayerTable cLayers = (LayerTable)tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead);

                // cicla sui layers e cerca i layers non stampabili
                foreach (ObjectId id in cLayers)
                {
                    LayerTableRecord lay = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);

                    if (!lay.IsPlottable && (lay.IsFrozen == bVisible))     // non stampabile e flag diverso da visibilità richiesta
                    {
                        lay.UpgradeOpen();

                        // se è il layer corrente, passa su layer "0" prima di spegnerlo
                        if (doc.Database.Clayer == id) doc.Database.Clayer = cLayers["0"];

                        if (lay.IsWriteEnabled) lay.IsFrozen = !bVisible;   // imposta visibilità richiesta
                    }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage("\n*** Errore inaspettato in [ChangeNonPrintableLayerVisibility]: " + ex.Message + " ***\n");
            }
            finally
            {
                tr.Dispose();
            }
        }

        public static void ChangeLayerVisibility(Document doc, String szLayerName, bool bVisible)
        {
            // esegui operazioni sul documento aperto
            Transaction tr = doc.TransactionManager.StartTransaction();

            try
            {
                // spegni/accendi layer 'Preliminare', se presente
                LayerTable cLayers = (LayerTable)tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead);
                if (cLayers.Has(szLayerName))
                {
                    // se è il layer corrente, passa su layer "0" prima di spegnerlo
                    if (doc.Database.Clayer == cLayers[szLayerName]) doc.Database.Clayer = cLayers["0"];

                    LayerTableRecord lay = (LayerTableRecord)tr.GetObject(cLayers[szLayerName], OpenMode.ForWrite);
                    if (lay.IsWriteEnabled) lay.IsFrozen = !bVisible;
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage("\n*** Errore inaspettato in [ChangeLayerVisibility]: " + ex.Message + " ***\n");
            }
            finally
            {
                tr.Dispose();
            }
        }

        static public PreviewEndPlotStatus _DoPlot(String szDocName, Database dwg, String szOutFolder, String szPlotterName, String szStyle, String szMedia, String szScale, bool bCentrapagina, bool bMultiPage)
        {
            PreviewEndPlotStatus result = PreviewEndPlotStatus.Normal;

            // controlli generici
            if (szPlotterName.Length == 0) return result;
            //if (szStyle.Length == 0) return result;

            short bgPlot = (short)Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("BACKGROUNDPLOT");
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("BACKGROUNDPLOT", 0);

            if (szMedia.StartsWith("auto", StringComparison.CurrentCultureIgnoreCase)) szMedia = "";         // forza selezione automatica del formato

            Point3d pMin, pMax;
            StdScaleType eScale = ConvertScaleToType(szScale);

            // apri transizione
            Transaction tr = dwg.TransactionManager.StartTransaction();

            try
            {
                // ricorda il layout corrente
                String szCurrentLayout = LayoutManager.Current.CurrentLayout;

                // recupera i layout da plottare
                ObjectIdCollection layoutsToPlot = GetLayoutIds(dwg);


                // The PlotSettingsValidator helps
                // create a valid PlotSettings object
                PlotSettingsValidator psv = PlotSettingsValidator.Current;


                if ((layoutsToPlot.Count > 0) && (szPlotterName.ToUpper().Contains(@"PDF")) && bMultiPage)
                {
                    // tenta stampa multipagina
                    int numSheet = 0;


                    // crea dialogo di avanzamento e consenti all'utente di annullare
                    PlotProgressDialog ppd = new PlotProgressDialog(false, layoutsToPlot.Count, true);

                    using (ppd)
                    {
                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Avanzamento stampa...");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetName, szDocName.Substring(szDocName.LastIndexOf("\\") + 1));
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Annulla Stampa");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Annulla pagina");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Avanzamento documento");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Avanzamento pagina");

                        // Inizia stampa/preview
                        ppd.OnBeginPlot();
                        ppd.IsVisible = true;

                        // crea engine per plottaggio diretto o in anteprima
                        PlotEngine pe = PlotFactory.CreatePublishEngine();

                        pe.BeginPlot(ppd, null);

                        // imposta cartella e nome file di uscita, se necessario
                        // stampa su file, probabilmente PDF o DWF
                        // controlla ultimo slash cartella uscita
                        if (!szOutFolder.EndsWith(@"\")) szOutFolder += @"\";

                        String szOutFile = szOutFolder + System.IO.Path.GetFileNameWithoutExtension(szDocName);

                        // trattamento casi speciali di stampanti PDF
                        String s = szPlotterName.ToUpper();

                        // determina estensione dal driver di stampa
                        if (s.Contains(@"PDF")) szOutFile += @".pdf";
                        else if (s.Contains(@"DWF")) szOutFile += @".dwf";
                        else szOutFile += @".plt";

                        // elimina pre-esistente se necessario
                        if (System.IO.File.Exists(szOutFile)) System.IO.File.Delete(szOutFile);


                        ppd.StatusMsgString = "Plotting " + System.IO.Path.GetFileNameWithoutExtension(szDocName);

                        ppd.OnBeginSheet();
                        ppd.LowerSheetProgressRange = 0;
                        ppd.UpperSheetProgressRange = 100;
                        ppd.SheetProgressPos = 0;

                        foreach (ObjectId o in layoutsToPlot)
                        {
                            numSheet++;
                            ppd.SheetProgressPos = numSheet * 100 / layoutsToPlot.Count;

                            // recupera layout
                            Layout lo = (Layout)tr.GetObject(o, OpenMode.ForRead);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(lo.BlockTableRecordId, OpenMode.ForRead);

                            pMin = lo.Extents.MinPoint;
                            pMax = lo.Extents.MaxPoint;

                            // rendi il layout corrente
                            LayoutManager.Current.CurrentLayout = lo.LayoutName;

                            // Inizializza configurazione di stampa specifica per il layout corrente
                            PlotSettings ps = new PlotSettings(lo.ModelType);
                            ps.CopyFrom(lo);

                            // apri un oggetto PlotInfo associato al layout corrente
                            PlotInfo pi = new PlotInfo();
                            pi.Layout = lo.ObjectId;

                            // assegna stile di stampa, se presente
                            _SetPlotStyle(szStyle, ps, psv);

                            ps.ScaleLineweights = false;
                            ps.PrintLineweights = true;
                            ps.PlotTransparency = true;

                            // Stampa le estensioni, con scala indicata e sempre centrata
                            psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                            psv.SetUseStandardScale(ps, true);
                            psv.SetStdScaleType(ps, eScale);
                            psv.SetPlotCentered(ps, bCentrapagina);
                            //psv.SetPlotPaperUnits(ps, PlotPaperUnit.Millimeters);

                            bool bPaperIsLandscape;
                            bool bAreaIsLandscape;

                            if (lo.ModelType) bAreaIsLandscape = (dwg.Extmax.X - dwg.Extmin.X) > (dwg.Extmax.Y - dwg.Extmin.Y) ? true : false;
                            else bAreaIsLandscape = (lo.Extents.MaxPoint.X - lo.Extents.MinPoint.X) > (lo.Extents.MaxPoint.Y - lo.Extents.MinPoint.Y) ? true : false;

                            // determina formato ideale
                            szMedia = _PlotSetMedia(szPlotterName, szMedia, lo, ps, psv);

                            // imposta rotazione automatica in base al formato ed all'area di stampa
                            bPaperIsLandscape = (ps.PlotPaperSize.X > ps.PlotPaperSize.Y) ? true : false;
                            if (bPaperIsLandscape != bAreaIsLandscape) psv.SetPlotRotation(ps, PlotRotation.Degrees270);

                            // collega a PlotInfo
                            pi.OverrideSettings = ps;

                            // valida parametri
                            PlotInfoValidator piv = new PlotInfoValidator();
                            piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                            piv.Validate(pi);

                            if (numSheet == 1)
                            {
                                // prima pagina, init documento

                                if (s.Contains(@"PDF") && s.Contains(@"ADOBE"))
                                {
                                    // stampante Adobe PDF
                                    // la stampa su file non funziona, ma genera un file postscript
                                    // occorre preimpostare una chiave di registro per indicare nome
                                    // e percorso del file di uscita per impedire la richiesta interattiva
                                    Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Adobe\Acrobat Distiller\PrinterJobControl",
                                                                      System.Windows.Forms.Application.ExecutablePath,
                                                                      szOutFile);

                                    // stampa diretta
                                    pe.BeginDocument(pi, szDocName, null, 1, false, null);
                                }
                                else
                                {
                                    // passa gestione file a driver AutoCAD.
                                    // Funziona bene per DWG_to_PDF e DWF, non funziona per altri driver PDF come PrimoPDF
                                    // perché genera un file Postscript
                                    pe.BeginDocument(pi, szDocName, null, 1, true, szOutFile);
                                }
                            }

                            if (pi.IsValidated)
                            {
                                // stampa pagina corrente
                                PlotPageInfo ppi = new PlotPageInfo();

                                pe.BeginPage(ppi, pi, (numSheet == layoutsToPlot.Count), null);
                                pe.BeginGenerateGraphics(null);

                                ppd.SheetProgressPos = 50;
                                pe.EndGenerateGraphics(null);

                                // Termina pagina
                                PreviewEndPlotInfo pepi = new PreviewEndPlotInfo();
                                pe.EndPage(pepi);
                                result = pepi.Status;
                                ppd.SheetProgressPos = 100;
                                ppd.OnEndSheet();
                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Termina documento
                        pe.EndDocument(null);

                        // Termina stampa
                        ppd.PlotProgressPos = 100;
                        ppd.OnEndPlot();
                        pe.EndPlot(null);
                        pe.Dispose();

                    }
                }
                else
                {
                    // alcune stampanti non sono in grado di gestire stampa multi-pagina
                    // tratta quindi ogni layout in stampa come stampa indipendente

                    int numSheet = 0;

                    foreach (ObjectId o in layoutsToPlot)
                    {
                        numSheet++;

                        // recupera layout
                        Layout lo = (Layout)tr.GetObject(o, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(lo.BlockTableRecordId, OpenMode.ForRead);

                        pMin = lo.Extents.MinPoint;
                        pMax = lo.Extents.MaxPoint;

                        // rendi il layout corrente
                        LayoutManager.Current.CurrentLayout = lo.LayoutName;

                        // Inizializza configurazione di stampa specifica per il layout corrente
                        PlotSettings ps = new PlotSettings(lo.ModelType);
                        ps.CopyFrom(lo);

                        // apri un oggetto PlotInfo associato al layout corrente
                        PlotInfo pi = new PlotInfo();
                        pi.Layout = btr.LayoutId;

                        // assegna stile di stampa, se presente
                        _SetPlotStyle(szStyle, ps, psv);

                        ps.ScaleLineweights = false;
                        ps.PrintLineweights = true;
                        ps.PlotTransparency = true;

                        // Stampa le estensioni, con scala indicata e sempre centrata
                        psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                        psv.SetUseStandardScale(ps, true);
                        psv.SetStdScaleType(ps, eScale);
                        psv.SetPlotCentered(ps, bCentrapagina);
                        //psv.SetPlotPaperUnits(ps, PlotPaperUnit.Millimeters);


                        bool bPaperIsLandscape;
                        bool bAreaIsLandscape;

                        if (lo.ModelType) bAreaIsLandscape = (dwg.Extmax.X - dwg.Extmin.X) > (dwg.Extmax.Y - dwg.Extmin.Y) ? true : false;
                        else bAreaIsLandscape = (lo.Extents.MaxPoint.X - lo.Extents.MinPoint.X) > (lo.Extents.MaxPoint.Y - lo.Extents.MinPoint.Y) ? true : false;

                        // determina formato ideale
                        szMedia = _PlotSetMedia(szPlotterName, szMedia, lo, ps, psv);

                        // imposta rotazione automatica in base al formato ed all'area di stampa
                        bPaperIsLandscape = (ps.PlotPaperSize.X > ps.PlotPaperSize.Y) ? true : false;
                        if (bPaperIsLandscape != bAreaIsLandscape) psv.SetPlotRotation(ps, PlotRotation.Degrees270);

                        // collega a PlotInfo
                        pi.OverrideSettings = ps;

                        // stampa pagina corrente
                        result = _DoPagePlot(szDocName, false, szOutFolder, szPlotterName, pi, numSheet, layoutsToPlot.Count, lo.LayoutName);

                        // attendi fine stampa prima di iniziarne un'altra
                        while (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting) System.Windows.Forms.Application.DoEvents();
                    }
                }

                // ripristina il layout precedente
                LayoutManager.Current.CurrentLayout = szCurrentLayout;

                //// commit più veloce di RollBack, ma fallisce se il disegno è readonly
                //tr.Commit();
            }
            catch (System.Exception ex)
            {
                result = PreviewEndPlotStatus.Cancel;
            }
            finally
            {
                tr.Dispose();
                Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);
            }

            return result;
        }
        public static StdScaleType ConvertScaleToType(String szScale)
        {
            StdScaleType eScale = StdScaleType.ScaleToFit;

            switch (szScale)
            {
                case "1:1":
                    eScale = StdScaleType.StdScale1To1;
                    break;

                case "1:2":
                    eScale = StdScaleType.StdScale1To2;
                    break;

                case "1:4":
                    eScale = StdScaleType.StdScale1To4;
                    break;

                case "1:8":
                    eScale = StdScaleType.StdScale1To8;
                    break;

                case "1:10":
                    eScale = StdScaleType.StdScale1To10;
                    break;

                case "1:20":
                    eScale = StdScaleType.StdScale1To20;
                    break;

                case "1:50":
                    eScale = StdScaleType.StdScale1To50;
                    break;

                case "1:100":
                    eScale = StdScaleType.StdScale1To100;
                    break;

                case "2:1":
                    eScale = StdScaleType.StdScale2To1;
                    break;

                case "4:1":
                    eScale = StdScaleType.StdScale4To1;
                    break;

                case "8:1":
                    eScale = StdScaleType.StdScale8To1;
                    break;

                case "10:1":
                    eScale = StdScaleType.StdScale10To1;
                    break;

                case "100:1":
                    eScale = StdScaleType.StdScale100To1;
                    break;
            }
            return eScale;
        }
        private static ObjectIdCollection GetLayoutIds(Database db)
        {
            ObjectIdCollection layoutIds = new ObjectIdCollection();

            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDic = Tx.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false) as DBDictionary;

                foreach (DBDictionaryEntry entry in layoutDic)
                {
                    Layout lay = (Layout)Tx.GetObject(entry.Value, OpenMode.ForRead);
                    if (lay.ModelType) continue;

                    if (lay.LayoutName.StartsWith("Layout"))
                    {
                        // verifica ci siano entità da stampare, ignora layout vuoti
                        BlockTableRecord btr = (BlockTableRecord)Tx.GetObject(lay.BlockTableRecordId, OpenMode.ForRead);

                        // verifica se il layout è valido, se è vuoto è inutile stampare
                        int count = 0;
                        foreach (ObjectId o in btr)
                        {
                            if (++count > 2) break;
                        }

                        // aggiungi solo se contiene qualche entità
                        // la prima entità è una Viewport
                        if (count < 2) continue;
                    }

                    layoutIds.Add(entry.Value);
                }

                // se non ci sono layout validi, aggiungi modelspace
                if (layoutIds.Count == 0)
                {
                    foreach (DBDictionaryEntry entry in layoutDic)
                    {
                        Layout lay = (Layout)Tx.GetObject(entry.Value, OpenMode.ForRead);
                        if (lay.ModelType)
                            layoutIds.Add(entry.Value);
                    }
                }
            }

            return layoutIds;
        }
        private static void _SetPlotStyle(String szStyle, PlotSettings ps, PlotSettingsValidator psv)
        {
            if (szStyle.Length > 0)
            {
                // recupera percorso completo dello stile di stampa
                System.Collections.Specialized.StringCollection sc = psv.GetPlotStyleSheetList();
                foreach (String szStyleFullName in sc)
                {
                    if (System.IO.Path.GetFileName(szStyleFullName) == szStyle)
                    {
                        psv.SetCurrentStyleSheet(ps, szStyleFullName);
                        break;
                    }
                }
            }
        }
        private static String _PlotSetMedia(String szPlotterName, String szMedia, Layout lo, PlotSettings ps, PlotSettingsValidator psv)
        {
            if (szMedia.Length > 0)
            {
                // imposta plotter corrente
                psv.SetPlotConfigurationName(ps, szPlotterName, null);
                psv.RefreshLists(ps);

                // formato indicato, usa
                // cerca nome canonico partendo dal nome suggerito
                System.Collections.Specialized.StringCollection v = psv.GetCanonicalMediaNameList(ps);
                String szCanonicalName = "", szTmp;


                foreach (String s in v)
                {
                    if (s == szMedia)
                    {
                        // il nome indicato è già un nome canonico, utilizza direttamente
                        szCanonicalName = szMedia;
                        break;
                    }
                    else
                    {
                        // ricava nome descrittivo locale e confronta
                        szTmp = psv.GetLocaleMediaName(ps, s);
                        if (szTmp == szMedia)
                        {
                            // ok, trovata corrispondenza
                            szCanonicalName = s;
                            break;
                        }
                    }
                }

                psv.SetPlotConfigurationName(ps, szPlotterName, szCanonicalName);       // stampante e formato pagina
                psv.SetPlotRotation(ps, PlotRotation.Degrees000);
            }
            else
            {
                // formato non indicato, cerca il migliore disponibile
                if (lo.ModelType) szMedia = FindBestFormat(ps, psv, szPlotterName, lo.Database.Extmin, lo.Database.Extmax);
                else szMedia = FindBestFormat(ps, psv, szPlotterName, lo.Extents.MinPoint, lo.Extents.MaxPoint);
            }
            return szMedia;
        }
        private static string FindBestFormat(PlotSettings ps, PlotSettingsValidator psv, string szPlotterName, Point3d pMin, Point3d pMax)
        {
            String result = "";
            String szLocalName, sztmp;
            bool bAreaIsLandscape = (pMax.X - pMin.X) > (pMax.Y - pMin.Y) ? true : false;

            // imposta plotter e forza ScaleToFit per forzare il calcolo della scana
            psv.SetPlotConfigurationName(ps, szPlotterName, null);   // stampante e formato pagina
            psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);

            // se è una conversione DWF il formato non serve
            if (!szPlotterName.ToUpper().Contains(@"DWF"))
            {
                System.Collections.Specialized.StringCollection vMedia = psv.GetCanonicalMediaNameList(ps);

                Double pDistX = pMax.X - pMin.X;
                Double pDistY = pMax.Y - pMin.Y;
                Double pDiagDist = Math.Sqrt(pDistX * pDistX + pDistY * pDistY);       // shorter reference

                Double dBestScale = 99999.9, dBestFit = 0.0;
                String szBestFormat = "", szBestFit = "";

                // cicla su tutti i formati disponibili per trovare quello più opportuno
                foreach (String sz in vMedia)
                {
                    szLocalName = psv.GetLocaleMediaName(ps, sz);
                    sztmp = szLocalName.ToLower();

                    // accetta solo formati ISO o definiti dall'utente
                    if (sztmp.Contains("a0") ||
                        sztmp.Contains("a1") ||
                        sztmp.Contains("a2") ||
                        sztmp.Contains("a3") ||
                        sztmp.Contains("a4") ||
                        sztmp.Contains("user"))
                    {
                        psv.SetPlotConfigurationName(ps, szPlotterName, sz);
                        psv.SetPlotRotation(ps, PlotRotation.Degrees000);

                        // determina rotazione ottimale per allineare i lati lunghi
                        bool bPaperIsLandscape = (ps.PlotPaperSize.X > ps.PlotPaperSize.Y) ? true : false;

                        Double dUseableX = ps.PlotPaperSize.X - ps.PlotPaperMargins.MinPoint.X - ps.PlotPaperMargins.MaxPoint.X;
                        Double dUseableY = ps.PlotPaperSize.Y - ps.PlotPaperMargins.MinPoint.Y - ps.PlotPaperMargins.MaxPoint.Y;

                        if (bPaperIsLandscape != bAreaIsLandscape) psv.SetPlotRotation(ps, PlotRotation.Degrees270);

                        // calcola scala sulla diagonale, in modo da avere una buona base
                        // sia per i formati orizzontali che verticali
                        Double dDiagScale = Math.Sqrt(dUseableX * dUseableX + dUseableY * dUseableY) / pDiagDist;

                        // la scala diagonale funziona bene se il rapporto tra i lati è coerente tra carta e modello
                        // altrimenti potrebbe debordare dal lato corto.
                        // Calcoliamo indipendentemente i fattori di scala sui lati per evitare fattori di scala inferiori
                        // ad 1 anche sui singoli lati
                        Double scale_x = ((bPaperIsLandscape == bAreaIsLandscape) ? dUseableX : dUseableY) / pDistX;
                        Double scale_y = ((bPaperIsLandscape == bAreaIsLandscape) ? dUseableY : dUseableX) / pDistY;

                        // ricorda scala migliore, ma superiore a 1
                        if ((dDiagScale > 1) && (scale_x > 1) && (scale_y > 1))
                        {
                            // trova migliore formato uguale o superiore al layout
                            if (dDiagScale < dBestScale)
                            {
                                dBestScale = dDiagScale;
                                szBestFormat = sz;
                            }
                        }
                        else
                        {
                            // trova migliore formato leggermente inferiore al layout richiesto
                            if (dDiagScale > dBestFit)
                            {
                                dBestFit = dDiagScale;
                                szBestFit = sz;
                            }
                        }
                    }
                }

                // decidi se utilizzare il formato minimo necessario oppure adattare un poco 
                // ed utilizzare quello leggermente inferiore
                if ((dBestScale < 1.10) || (szBestFit.Length == 0))
                {
                    // spreca un po' di carta, ma utilizza scala 1:1
                    psv.SetPlotConfigurationName(ps, szPlotterName, szBestFormat);
                    psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                    psv.SetStdScaleType(ps, StdScaleType.StdScale1To1);
                    result = szBestFormat;
                }
                else
                {
                    // utilizza formato inferiore con ScaleToFit
                    psv.SetPlotConfigurationName(ps, szPlotterName, szBestFit);
                    psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                    psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                    result = szBestFit;
                }
            }

            return result;
        }
        private static PreviewEndPlotStatus _DoPagePlot(String szDocName, bool bPreview, String szOutFolder, String szPlotterName, PlotInfo pi, int numSheet, int numTot, String szSuffix)
        {
            PreviewEndPlotStatus result = PreviewEndPlotStatus.Normal;

            try
            {
                // valida parametri
                PlotInfoValidator piv = new PlotInfoValidator();
                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                piv.Validate(pi);

                // crea dialogo di avanzamento e consenti all'utente di annullare
                PlotProgressDialog ppd = new PlotProgressDialog(bPreview, 1, true);

                using (ppd)
                {
                    ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Avanzamento stampa...");
                    ppd.set_PlotMsgString(PlotMessageIndex.SheetName, szDocName.Substring(szDocName.LastIndexOf("\\") + 1));
                    ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Annulla Stampa");
                    ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Annulla pagina");
                    ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Avanzamento documento");
                    ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Avanzamento pagina");

                    // Inizia stampa/preview
                    ppd.OnBeginPlot();
                    ppd.IsVisible = false;

                    // crea engine per plottaggio diretto o in anteprima
                    PlotEngine pe;

                    if (bPreview)
                        // inizializza un controllo Preview
                        pe = PlotFactory.CreatePreviewEngine((int)PreviewEngineFlags.Plot);
                    else
                        pe = PlotFactory.CreatePublishEngine();

                    pe.BeginPlot(ppd, null);

                    // imposta cartella e nome file di uscita, se necessario
                    _PlotSetOutput(szDocName, szOutFolder, szPlotterName, pe, pi, numSheet, numTot, szSuffix);

                    ppd.StatusMsgString = "Plotting " + szDocName.Substring(szDocName.LastIndexOf("\\") + 1);

                    ppd.OnBeginSheet();
                    ppd.LowerSheetProgressRange = 0;
                    ppd.UpperSheetProgressRange = 100;
                    ppd.SheetProgressPos = 0;

                    PlotPageInfo ppi = new PlotPageInfo();

                    pe.BeginPage(ppi, pi, true, null);
                    pe.BeginGenerateGraphics(null);

                    ppd.SheetProgressPos = 50;
                    pe.EndGenerateGraphics(null);

                    // Termina pagina
                    PreviewEndPlotInfo pepi = new PreviewEndPlotInfo();
                    pe.EndPage(pepi);
                    result = pepi.Status;
                    ppd.SheetProgressPos = 100;
                    ppd.OnEndSheet();

                    //System.Windows.Forms.Application.DoEvents();

                    // Termina documento
                    pe.EndDocument(null);

                    // Termina stampa
                    ppd.PlotProgressPos = 100;
                    ppd.OnEndPlot();
                    pe.EndPlot(null);
                    pe.Dispose();
                }
            }
            catch
            {
                // errore durante stampa, annulla
                result = PreviewEndPlotStatus.Cancel;
            }

            return result;
        }
        private static void _PlotSetOutput(String docName, String szOutFolder, String szPlotterName, PlotEngine pe, PlotInfo pi, int numSheet, int numTot, String szSuffix)
        {
            if (szOutFolder.Length == 0)
            {
                // stampa diretta a stampante
                pe.BeginDocument(pi, docName, null, 1, false, null);
            }
            else
            {
                // stampa su file, probabilmente PDF o DWF
                // controlla ultimo slash cartella uscita
                if (!szOutFolder.EndsWith(@"\")) szOutFolder += @"\";

                String szOutFile = szOutFolder + System.IO.Path.GetFileNameWithoutExtension(docName);

                // in caso di stampa di layout multipli aggiungi suffisso
                //if (szSuffix.Length > 0) szOutFile += "_" + szSuffix;
                //else if (numTot > 1) szOutFile += "_Page_" + numSheet.ToString() + "_of_" + numTot.ToString();

                // trattamento casi speciali di stampanti PDF
                String s = szPlotterName.ToUpper();

                // determina estensione dal driver di stampa
                if (s.Contains(@"PDF")) szOutFile += @".pdf";
                else if (s.Contains(@"DWF")) szOutFile += @".dwf";
                else szOutFile += @".plt";

                // elimina pre-esistente se necessario
                if (System.IO.File.Exists(szOutFile)) System.IO.File.Delete(szOutFile);

                if (s.Contains(@"PDF") && s.Contains(@"ADOBE"))
                {
                    // stampante Adobe PDF
                    // la stampa su file non funziona, ma genera un file postscript
                    // occorre preimpostare una chiave di registro per indicare nome
                    // e percorso del file di uscita per impedire la richiesta interattiva
                    Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Adobe\Acrobat Distiller\PrinterJobControl",
                                                      System.Windows.Forms.Application.ExecutablePath,
                                                      szOutFile);

                    // stampa diretta
                    pe.BeginDocument(pi, docName, null, 1, false, null);
                }
                //else if (s.Contains(@"BULLZIP"))
                //{
                //    Bullzip.PDFPrinterSettingsClass pdfSetting = new Bullzip.PDFPrinterSettingsClass();
                //    pdfSetting.Init();
                //    pdfSetting.SetPrinterName(szPlotterName);
                //    pdfSetting.SetValue(@"output", szOutFile);
                //    pdfSetting.SetValue(@"showpdf", @"no");
                //    pdfSetting.SetValue(@"showsettings", @"never");
                //    pdfSetting.SetValue(@"showsaveas", @"never");
                //    pdfSetting.SetValue(@"target", @"ebook");
                //    pdfSetting.WriteSettings(true);

                //    // stampa diretta
                //    pe.BeginDocument(pi, docName, null, 1, false, null);
                //}
                else
                {
                    // passa gestione file a driver AutoCAD.
                    // Funziona bene per DWG_to_PDF e DWF, non funziona per altri driver PDF come PrimoPDF
                    // perché genera un file Postscript
                    pe.BeginDocument(pi, docName, null, 1, true, szOutFile);
                }
            }
        }

    }
    public class SimpleButtonCmdHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object param)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object param)
        {
            if (param is RibbonButton)
            {
                String esc = "";

                // se un comando è attivo, costruisci sequenza di escapes da anteporre al comando
                String cmds = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDNAMES") as String;

                if (cmds.Length > 0)
                {
                    //esc = "\x1b\x1b\x1b";
                    esc = "\x03\x03\x03";
                }

                RibbonButton btn = param as RibbonButton;
                Document doc = Application.DocumentManager.MdiActiveDocument;

                if (doc != null)
                    doc.SendStringToExecute(esc + btn.CommandParameter, true, false, false);
            }
        }
    }
}
