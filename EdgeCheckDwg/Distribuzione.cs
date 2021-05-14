using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace EdgeAutocadPlugins
{
    class Distribuzione
    {

        private string ip = "192.168.100.2";
        private string dbName = "bim_200000";
        private string username = "bim";
        private string password = "Quink$2100";

        ///SELECT 
        ///t.ha_filename, c.ad_figlio AS code, Sum(ac_qty) AS totale
        ///FROM (((((analisi_diba c LEFT JOIN analisi_diba up on up.ad_figlio = c.ad_padre)  
        ///LEFT JOIN analisi_diba upp on upp.ad_figlio = up.ad_padre)  
        ///LEFT JOIN analisi_diba uppp on uppp.ad_figlio = upp.ad_padre)  
        ///LEFT JOIN analisi_codes d ON d.ac_code = COALESCE(uppp.ad_padre, upp.ad_padre, up.ad_padre, c.ad_padre))  
        ///LEFT JOIN analisi_history t on t.id = ac_analisi_id)  
        ///WHERE (t.ha_exclude = 0)  
        ///GROUP BY c.ad_figlio, t.ha_filename ORDER BY c.ad_figlio;
        
        //ha_filename      code            totale
        //10009	        GU2116-001	    2
        public void queryFocchiStartup()
        {
            // Da qui ottengo le analisi attive
            DataTable analisi_history = this.call("SELECT * FROM bim_200000.analisi_history WHERE ha_exclude = 0;");

            int[] analisiAttive = (from DataRow dr in analisi_history.Rows select (int)dr["id"]).ToArray();
            string[] codiciModello = (from DataRow dr in analisi_history.Rows select (string)dr["ha_filename"]).ToArray();

            Dictionary<int, string> v = new Dictionary<int, string>();
            
            int index = 0;
            foreach (int an in analisiAttive)
            {
                v[an] = codiciModello[index];
                index += 1;
            }


            // Prendo solo da analisi attive
            DataTable analisi_codes = this.call("SELECT * FROM bim_200000.analisi_codes where ac_analisi_id in (" + string.Join(",", analisiAttive.Select(x => x.ToString()).ToArray()) + ");");

            List<Obj> cellules = new List<Obj>();

            foreach (DataRow dr in analisi_codes.Rows)
            {
                string ac_code = (string)dr["ac_code"];
                int ac_analisi_id = (int)dr["ac_analisi_id"];
                int ac_qty = (int)dr["ac_qty"];


                Obj std = new Obj()
                {
                    obj = ac_code,
                    qty = ac_qty,
                    file = (v.ContainsKey(ac_analisi_id) ? v[ac_analisi_id] : null)
                };

                cellules.Add(std);
            }

            DataTable analisi_diba = this.call("SELECT * FROM bim_200000.analisi_diba;");

            IEnumerable<string> listCellule = cellules.Select(i => i.obj).ToList().Distinct();
            Dictionary<string, List<string>> padrefiglio = new Dictionary<string, List<string>>();
            
            giaAnalizzato = new List<string>();

            foreach (string c in listCellule)
            {
                List<string> r = ric(c, analisi_diba);

                //IList<string> uniquer = r.Distinct().ToList();
                List<string> uniquer = r.ToList();

                padrefiglio.Add(c, uniquer);
            }


            //IEnumerable<string> listFile = cellules.Select(i => i.file).ToList().Distinct();

            Dictionary<string, Dictionary<string, int>> lottoCellQuant = new Dictionary<string, Dictionary<string, int>>();

            foreach (Obj o in cellules)
            {
                string nomeFile = o.file;
                string obj = o.obj;

                if (!lottoCellQuant.ContainsKey(nomeFile))
                {
                    lottoCellQuant.Add(nomeFile, new Dictionary<string, int>());
                }

                if (!lottoCellQuant[nomeFile].ContainsKey(obj))
                {
                    lottoCellQuant[nomeFile].Add(obj, 0);
                }

                lottoCellQuant[nomeFile][obj] += 1;
            }

            Dictionary<string, Dictionary<string, int>> results = new Dictionary<string, Dictionary<string, int>>();

            foreach (KeyValuePair<string, Dictionary<string, int>> d in lottoCellQuant)
            {
                string nomeFile = d.Key;

                if (!results.ContainsKey(nomeFile)) 
                    results.Add(nomeFile, new Dictionary<string, int>());

                Dictionary<string, int> value = d.Value;

                foreach (KeyValuePair<string, int> c in value)
                {
                    string codiceCellula = c.Key;
                    int quantita = c.Value;

                    if (padrefiglio.ContainsKey(codiceCellula))
                    {
                        List<string> dist = padrefiglio[codiceCellula];
                        var q = from x in dist
                                group x by x into g
                                let count = g.Count()
                                orderby count descending
                                select new { Value = g.Key, Count = count };

                        foreach (var x in q)
                        {
                            string code = x.Value;
                            int qta = quantita * x.Count;

                            if(!results[nomeFile].ContainsKey(code))
                                    results[nomeFile].Add(code,0);

                            results[nomeFile][code] += qta;
                            //DataRow _ravi = dt.NewRow();
                            //_ravi["ha_filename"] = nomeFile;
                            //_ravi["code"] = code;
                            //_ravi["totale"] = quantita;
                            //dt.Rows.Add(_ravi);
                        }
                    }
                }
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("ha_filename");
            dt.Columns.Add("code");
            dt.Columns.Add("totale");


            foreach (KeyValuePair<string, Dictionary<string, int>> d in results)
            {
                string nomeFile = d.Key;
                Dictionary<string, int> value = d.Value;

                foreach (KeyValuePair<string, int> c in value)
                {
                    string codice = c.Key;
                    int quantita = c.Value;


                    DataRow _ravi = dt.NewRow();
                    _ravi["ha_filename"] = nomeFile;
                    _ravi["code"] = codice;
                    _ravi["totale"] = quantita;
                    dt.Rows.Add(_ravi);
                }
            }

            //dt = dt.AsEnumerable()
            //   .GroupBy(r => new { Col1 = r["ha_filename"], Col2 = r["code"] })
            //   .Select(g => g.OrderBy(r => r["code"]).First())
            //   .CopyToDataTable();


            //DataRow[] a = dt.Select("ha_filename LIKE '10002'");

            //foreach (DataRow i in a)
            //{
            //    Console.WriteLine(i["code"] + "," + i["totale"]);
            //}

            Console.WriteLine(dt.Rows.Count);
        }



        /// SELECT
        ///c.ad_figlio AS codice, d.ac_lot AS id
        ///FROM
        ///((((analisi_diba c
        ///LEFT JOIN analisi_diba up ON up.ad_figlio = c.ad_padre)
        ///LEFT JOIN analisi_diba upp ON upp.ad_figlio = up.ad_padre)
        ///LEFT JOIN analisi_codes d ON d.ac_code = COALESCE(upp.ad_padre, up.ad_padre, c.ad_padre))
        ///LEFT JOIN analisi_history t ON t.id = ac_analisi_id)
        ///WHERE
        ///(t.ha_exclude = 0)
        ///GROUP BY c.ad_figlio , d.ac_lot;

        //codice       id
        //GU1008-003	CC01C
        public void queryFocchiDesign()
        {

            // Da qui ottengo le analisi attive
            DataTable analisi_history = this.call("SELECT * FROM bim_200000.analisi_history WHERE ha_exclude = 0;");

            int[] analisiAttive = (from DataRow dr in analisi_history.Rows select (int)dr["id"]).ToArray();

            // Prendo solo da analisi attive
            DataTable analisi_codes = this.call("SELECT * FROM bim_200000.analisi_codes where ac_analisi_id in (" + string.Join(",", analisiAttive.Select(x => x.ToString()).ToArray()) + ");");

            // lotid - [cellula]
            //Dictionary<string, IList<string>> distr_old = new Dictionary<string, IList<string>>();
            Dictionary<string, List<string>> distr = new Dictionary<string, List<string>>();

            foreach (DataRow dr in analisi_codes.Rows)
            {
                string ac_code = (string)dr["ac_code"];
                string ac_lot = (string)dr["ac_lot"];

                if (!distr.ContainsKey(ac_lot))
                {
                    distr.Add(ac_lot, new List<string>());
                }

                if (!distr[ac_lot].Contains(ac_code))
                {
                    distr[ac_lot].Add(ac_code);
                }

            }

            // Qui funzione ricorsiva
            DataTable analisi_diba = this.call("SELECT * FROM bim_200000.analisi_diba;");

            List<string> allCeCodes = new List<string>();

            foreach (List<string> l in distr.Values)
            {
                foreach (string s in l)
                {
                    if (!allCeCodes.Contains(s))
                    {
                        allCeCodes.Add(s);
                    }
                }
            }


            // Cerco le distinte dei codici cellula

            //IDictionary<string, IList<string>> padrefiglio_old = new Dictionary<string, IList<string>>();
            Dictionary<string, List<string>> padrefiglio = new Dictionary<string, List<string>>();

            giaAnalizzato = new List<string>();

            foreach (string c in allCeCodes)
            {
                List<string> r = ric(c, analisi_diba);

                //IList<string> uniquer = r.Distinct().ToList();
                List<string> uniquer = r.Distinct().ToList();

                padrefiglio.Add(c, uniquer);
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("codice");
            dt.Columns.Add("id");

            foreach (KeyValuePair<string, List<string>> v in distr)
            {
                string lotid = v.Key;
                List<string> ces = v.Value;

                foreach (string c in ces)
                {
                    if (padrefiglio.ContainsKey(c))
                    {
                        foreach (string o in padrefiglio[c])
                        {
                            DataRow _ravi = dt.NewRow();
                            _ravi["codice"] = o;
                            _ravi["id"] = lotid;
                            dt.Rows.Add(_ravi);
                        }
                    }
                }
            }

            dt = dt.AsEnumerable()
               .GroupBy(r => new { Col1 = r["codice"], Col2 = r["id"] })
               .Select(g => g.OrderBy(r => r["codice"]).First())
               .CopyToDataTable();

            //DataRow[] a = dt.Select("id LIKE 'CB01C'");

            //foreach (DataRow i in a)
            //{
            //    Console.WriteLine(i["codice"]+","+i["id"]);
            //}

            Console.WriteLine(dt.Rows.Count);
        }



        ///SELECT 
        ///    c.ad_figlio AS codice,
        ///    ac_lot,
        ///    ac_preorder,
        ///    ac_level,
        ///    SUM(ac_qty) AS totale
        ///FROM
        ///    ((((analisi_diba c
        ///    LEFT JOIN analisi_diba up ON up.ad_figlio = c.ad_padre)
        ///    LEFT JOIN analisi_diba upp ON upp.ad_figlio = up.ad_padre)
        ///    LEFT JOIN analisi_codes d ON d.ac_code = COALESCE(upp.ad_padre, up.ad_padre, c.ad_padre))
        ///    LEFT JOIN analisi_history t ON t.id = ac_analisi_id)
        ///WHERE
        ///    (t.ha_exclude = 0)
        ///GROUP BY c.ad_figlio , d.ac_lot , d.ac_preorder , d.ac_level;

        // codice       ac_lot      ac_preorder     ac_level    totale
        // GU1008-003	CC01C	    CGF	            02	        1

        public void queryFocchiDesign2()
        {

            // Da qui ottengo le analisi attive
            DataTable analisi_history = this.call("SELECT * FROM bim_200000.analisi_history WHERE ha_exclude = 0;");

            int[] analisiAttive = (from DataRow dr in analisi_history.Rows select (int)dr["id"]).ToArray();

            // Prendo solo da analisi attive
            DataTable analisi_codes = this.call("SELECT * FROM bim_200000.analisi_codes where ac_analisi_id in (" + string.Join(",", analisiAttive.Select(x => x.ToString()).ToArray()) + ");");

            // lotid - [cellula]
            Dictionary<string, List<string>> distr = new Dictionary<string, List<string>>();

            foreach (DataRow dr in analisi_codes.Rows)
            {
                string ac_code = (string)dr["ac_code"];
                string ac_lot = (string)dr["ac_lot"];

                if (!distr.ContainsKey(ac_lot))
                {
                    distr.Add(ac_lot, new List<string>());
                }

                if (!distr[ac_lot].Contains(ac_code))
                {
                    distr[ac_lot].Add(ac_code);
                }

            }

            // Qui funzione ricorsiva
            DataTable analisi_diba = this.call("SELECT * FROM bim_200000.analisi_diba;");

            List<string> allCeCodes = new List<string>();

            foreach (List<string> l in distr.Values)
            {
                foreach (string s in l)
                {
                    if (!allCeCodes.Contains(s))
                    {
                        allCeCodes.Add(s);
                    }
                }
            }


            // Cerco le distinte dei codici cellula

            Dictionary<string, List<string>> padrefiglio = new Dictionary<string, List<string>>();

            giaAnalizzato = new List<string>();

            foreach (string c in allCeCodes)
            {
                List<string> r = ric(c, analisi_diba);

                List<string> uniquer = r.Distinct().ToList();

                padrefiglio.Add(c, uniquer);
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("codice");
            dt.Columns.Add("id");

            foreach (KeyValuePair<string, List<string>> v in distr)
            {
                string lotid = v.Key;
                List<string> ces = v.Value;

                foreach (string c in ces)
                {
                    if (padrefiglio.ContainsKey(c))
                    {
                        foreach (string o in padrefiglio[c])
                        {
                            DataRow _ravi = dt.NewRow();
                            _ravi["codice"] = o;
                            _ravi["id"] = lotid;
                            dt.Rows.Add(_ravi);
                        }
                    }
                }
            }

            dt = dt.AsEnumerable()
               .GroupBy(r => new { Col1 = r["codice"], Col2 = r["id"] })
               .Select(g => g.OrderBy(r => r["codice"]).First())
               .CopyToDataTable();

            //DataRow[] a = dt.Select("id LIKE 'CB01C'");

            //foreach (DataRow i in a)
            //{
            //    Console.WriteLine(i["codice"]+","+i["id"]);
            //}

            Console.WriteLine(dt.Rows.Count);
        }



        public Dictionary<string, dynamic> customQueryLotID(List<string> codiciAmmessi)
        {

            // Da qui ottengo le analisi attive
            DataTable analisi_history = this.call("SELECT * FROM bim_200000.analisi_history WHERE ha_exclude = 0;");

            int[] analisiAttive = (from DataRow dr in analisi_history.Rows select (int)dr["id"]).ToArray();

            // Prendo solo da analisi attive
            DataTable analisi_codes = this.call("SELECT * FROM bim_200000.analisi_codes where ac_analisi_id in (" + string.Join(",", analisiAttive.Select(x => x.ToString()).ToArray()) + ");");

            // lotid - [cellula]
            Dictionary<string, List<string>> distr = new Dictionary<string, List<string>>();

            foreach (DataRow dr in analisi_codes.Rows)
            {
                string ac_code = (string)dr["ac_code"];
                string ac_lot = (string)dr["ac_lot"];

                if (codiciAmmessi.Contains(ac_code))
                {
                    if (!distr.ContainsKey(ac_lot))
                    {
                        distr.Add(ac_lot, new List<string>());
                    }

                    distr[ac_lot].Add(ac_code);
                }
            }

            Dictionary<string, dynamic> conteggio = new Dictionary<string, dynamic>();

            foreach (KeyValuePair<string, List<string>> kv in distr)
            {
                string lotto = kv.Key;

                var q = from x in kv.Value
                        group x by x into g
                        let count = g.Count()
                        orderby count descending
                        select new { Value = g.Key, Count = count };

                conteggio.Add(lotto, q);
            }

            return conteggio;
        }

        public Dictionary<string, dynamic> customQueryFasciePiani(List<string> codiciAmmessi)
        {

            // Da qui ottengo le analisi attive
            DataTable analisi_history = this.call("SELECT * FROM bim_200000.analisi_history WHERE ha_exclude = 0;");

            int[] analisiAttive = (from DataRow dr in analisi_history.Rows select (int)dr["id"]).ToArray();

            // Prendo solo da analisi attive
            DataTable analisi_codes = this.call("SELECT * FROM bim_200000.analisi_codes where ac_analisi_id in (" + string.Join(",", analisiAttive.Select(x => x.ToString()).ToArray()) + ");");

            // lotid - [cellula]
            Dictionary<string, List<string>> distr = new Dictionary<string, List<string>>();

            foreach (DataRow dr in analisi_codes.Rows)
            {
                string ac_code = (string)dr["ac_code"];
                string ac_preorder = (string)dr["ac_preorder"];

                if (codiciAmmessi.Contains(ac_code))
                {
                    if (!distr.ContainsKey(ac_preorder))
                    {
                        distr.Add(ac_preorder, new List<string>());
                    }

                    distr[ac_preorder].Add(ac_code);
                }
            }

            Dictionary<string, dynamic> conteggio = new Dictionary<string, dynamic>();

            foreach (KeyValuePair<string, List<string>> kv in distr)
            {
                string lotto = kv.Key;

                var q = from x in kv.Value
                        group x by x into g
                        let count = g.Count()
                        orderby count descending
                        select new { Value = g.Key, Count = count };

                conteggio.Add(lotto, q);
            }

            return conteggio;
        }








        public DataTable call(string query)
        {
            DataTable dt = new DataTable();

            string constring = "server= " + this.ip + "; database =" + this.dbName + "; username= " + this.username + "; password=" + this.password + ";";

            using (MySqlConnection conn = new MySqlConnection(constring))
            {
                using (MySqlCommand cmd = new MySqlCommand())
                {
                    cmd.Connection = conn;
                    conn.Open();

                    cmd.CommandText = query;
                    dt.Load(cmd.ExecuteReader());
                }
            }

            return dt;
        }

        public List<string> giaAnalizzato = new List<string>();

        public List<string> ric(string code, DataTable analisi_diba)
        {
            List<string> l = new List<string>();

            IEnumerable<DataRow> r = analisi_diba.Rows.Cast<DataRow>().Where(x => x.Field<string>("ad_padre") == code);

            foreach (DataRow dr in r)
            {
                string f = dr.Field<string>("ad_figlio");

                l.Add(f);

                string name = cleanObj(f);

                if (!giaAnalizzato.Contains(cleanObj(name)))
                {
                    giaAnalizzato.Add(name);
                    l = l.Concat(ric(f, analisi_diba)).ToList();
                }
            }

            return l;
        }

        public string cleanObj(string s)
        {
            return s.Split('-')[0];
        }

        internal class Obj
        {
            public string obj { get; set; }
            public int qty { get; set; }
            public string file { get; set; }
        }
    }
}