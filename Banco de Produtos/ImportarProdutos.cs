using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Banco_de_Produtos
{
    public partial class ImportarProdutos : Form
    {
        public ImportarProdutos()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();

            var archive = openFileDialog1.FileName;

            if (File.Exists(archive))
            {
                textBox1.Text = archive;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            StringBuilder text = new StringBuilder();

            PdfReader pdfReader = new PdfReader(textBox1.Text);

            for (int page = 1; page <= pdfReader.NumberOfPages; page++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                currentText = currentText.Replace("\n", "\r\n").Replace("=\r\n", "").Replace("=", "");
                var textBreak = "Codigo e Descricao do Produto --------------------*  $Compra  $Varejo   $Atac.  $Difer. QntRepos\r\n";
                currentText = currentText.Substring(currentText.IndexOf(textBreak) + textBreak.Length);

                currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                text.Append(currentText);
            }

            pdfReader.Close();

            this.textBox2.Invoke(new MethodInvoker(delegate () { this.textBox2.Text = text.ToString(); }));
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            backgroundWorkerImport.RunWorkerAsync();
        }

        private void backgroundWorkerImport_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] lines = textBox2.Text.Replace("\r\n", "\n").Split('\n');
            this.progressBar1.Invoke(new MethodInvoker(delegate () { this.progressBar1.Maximum = lines.Length; }));

            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|/Database1.mdb");

            OleDbCommand cmd = con.CreateCommand();
            con.Open();
            cmd.Connection = con;

            int executed = 0;
            foreach(String line in lines)
            {
                executed++;
                backgroundWorkerImport.ReportProgress(executed);

                string codigo = "";
                
                if(line.Contains(" "))
                {

                    if (int.TryParse(line.Substring(0, line.IndexOf(" ")), out int result))
                    {
                        codigo = line.Substring(0, line.IndexOf(" "));
                    }
                }

                if(codigo == "" || line.Contains("Produto(s) neste Grupo") || line.Contains("Produto(s) no Total"))
                {
                    if(line.Contains("Grupo"))
                        MessageBox.Show("Categoria: " + line);
                }
                else
                {
                    String descricao = "";

                    try
                    {
                        descricao = line.Substring(codigo.Length + 1, line.IndexOf("   ") - (codigo.Length + 1)).Replace("'", "_");
                    }catch
                    {
                        Console.WriteLine(line);
                        Application.Exit();
                    }

                    String startCharacter = line.Substring(descricao.Length +codigo.Length + 3);

                    string[] teste = startCharacter.Split(' ');

                    double compra = 0;
                    double varejo = 0;
                    double atacado = 0;
                    double diferenca = 0;
                    double quantidade = 0;

                    int count = 0;
                    foreach(String a in teste)
                    {
                        if(double.TryParse(a, out double ab))
                        {
                            count++;
                            switch (count)
                            {
                                case 1:
                                    compra = ab;
                                    break;
                                case 2:
                                    varejo = ab;
                                    break;
                                case 3:
                                    atacado = ab;
                                    break;
                                case 4:
                                    diferenca = ab;
                                    break;
                                case 5:
                                    quantidade = ab;
                                    break;
                            }
                        }
                    }

                    cmd.CommandText = "Insert into Produtos(Codigo,Descricao, Compra, Varejo, Atacado, Diferenca, Quantidade) Values('" + codigo + "','" + descricao + "', '"+ compra +"', '"+varejo+"', '"+ atacado +"', '"+ diferenca +"', '"+ quantidade +"')";
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("linha gravada: " + executed);
                }
            }

            con.Close();
        }

        private void backgroundWorkerImport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = false;
            MessageBox.Show("Importação finalizada");
        }

        private void backgroundWorkerImport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
    }
}
