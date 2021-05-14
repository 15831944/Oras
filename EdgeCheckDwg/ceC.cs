using System.Data;

namespace EdgeAutocadPlugins
{
    internal class ceC
    {
        internal string TYPE;

        public string DESCRIZIONE { get; internal set; }
        public string CLASSE { get; internal set; }
        public string SOTTOCLASSE { get; internal set; }
        public string EMESSO { get; internal set; }
        public string COMMESSA { get; internal set; }
        public string NOME { get; internal set; }
        public string NOMEDWG { get; internal set; }
        public string NOMEPDF { get; internal set; }
        public DataTable POSIZIONI { get; internal set; }
        public DataTable DIBA { get; internal set; }
        public string NOMEEMI { get; internal set; }
    }
}