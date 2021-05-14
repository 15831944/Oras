using System;
using System.Data;
using System.Windows.Forms;

namespace EdgeAutocadPlugins
{
    public partial class ConfrontoDwg : Form
    {
        public static DataGridView dw1;
        public static DataGridView dw2;

        public ConfrontoDwg()
        {
            InitializeComponent();
        }

        private void ConfrontoDwg_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = GenericFunction.choosePath();
            EdgeAutocadPlugins.Commands.folderRecenti = path;
            label3.Text = path;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = GenericFunction.choosePath();
            EdgeAutocadPlugins.Commands.folderInviati = path;
            label4.Text = path;
        }

        private void c1Button1_Click(object sender, EventArgs e)
        {
            EdgeAutocadPlugins.Commands.avviaCheckPE(this);
        }
    }
}
