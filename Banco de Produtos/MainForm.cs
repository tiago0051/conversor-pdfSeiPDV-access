using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Banco_de_Produtos
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {
            var importarProdutos = new ImportarProdutos();
            importarProdutos.ShowDialog();
            backgroundWorkerUpdateList.RunWorkerAsync();
        }

        public void updateList()
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|/Database1.mdb");

            OleDbCommand cmd = con.CreateCommand();
            con.Open();
            cmd.Connection = con;

            cmd.CommandText = "SELECT * FROM produtos";

            OleDbDataReader reader = cmd.ExecuteReader();

            List<ListViewItem> listItems = new List<ListViewItem>();

            while (reader.Read())
            {
                String descricao = reader.GetString(1);
                Double varejo = double.Parse(reader.GetValue(2).ToString());
                Double atacado = double.Parse(reader.GetValue(3).ToString());

                ListViewItem item = new ListViewItem(new[] { descricao, varejo.ToString("C"), atacado.ToString("C") });
                listItems.Add(item);
            }

            Console.WriteLine("Executed: " + listItems.Count);

            this.listView1.Invoke(new MethodInvoker(delegate () {

                this.listView1.Items.AddRange(listItems.ToArray());
             }));

            reader.Close();
            con.Close();
        }

        private void backgroundWorkerUpdateList_DoWork(object sender, DoWorkEventArgs e)
        {
            updateList();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            backgroundWorkerUpdateList.RunWorkerAsync();
        }
    }
}
