using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Text;
using ClosedXML.Excel;

namespace Antifurto_Relatorio
{
    public partial class Form1 : Form
    {
        string conStr = "User Id=rmsprod;Password=rmsprod;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.1.14)(PORT=1521)))(CONNECT_DATA=(SID=sandbox)))";

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true; // Habilitar a captura de teclas no formulário
            this.KeyDown += Form1_KeyDown; // Associar o evento KeyDown ao manipulador de evento
        }
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)// Chamar o evento do botão executar
            {
                button2_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F3)// Chamar o evento do botão sair
            {
                button1_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F6)// Chamar o evento do botão converte texto
            {
                button3_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F10)// Chamar o evento do botão excel
            {
                button4_Click(sender, e);
            }
        }
        private void consulta()
        {
            string sqlQuery = "select det_cod_item||'-'||dac(det_cod_item) as cod_item " +
                "\r\n,git_descricao as descrição" +
                "\r\n,DET_ANTI_FURTO  as Anti_furto" +
                "\r\nfrom aa1ditem, AA3CITEM " +
                "\r\nwhere DET_ANTI_FURTO = 'S'" +
                "\r\nand det_cod_item = git_cod_item\r\n";

            using (OracleConnection con = new OracleConnection(conStr))
            {
                con.Open();

                using (OracleCommand command = new OracleCommand(sqlQuery, con))
                {
                    
                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        dataGridView1.DataSource = dt; // Configurar o DataGridView para exibir os dados
                    }
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();//fechar o formulario
        }

        private void button2_Click(object sender, EventArgs e)
        {
            consulta();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportTotxt();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }
        private void ExportTotxt()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Não há dados para exportar.");
                return;
            }
            // Abrir janela de diálogo para escolher onde salvar o arquivo
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Relatorio de itens com antifurto (*.txt)|*.txt";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Escrever os dados no arquivo CSV
                    using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                    {
                        // Escrever o cabeçalho
                        writer.WriteLine(string.Join(",", dataGridView1.Columns.Cast<DataGridViewColumn>().Select(column => column.HeaderText)));

                        // Escrever as linhas de dados
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            writer.WriteLine(string.Join(",", row.Cells.Cast<DataGridViewCell>().Select(cell => cell.Value)));
                        }
                    }

                    MessageBox.Show("Dados exportados para TXT com sucesso!");
                }
            }
        }

        private void ExportToExcel()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Não há dados para exportar.");
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // Defina o nome de arquivo predefinido aqui
                saveFileDialog.FileName = "Relatorio de itens com antifurto.xlsx"; // Altere o nome conforme necessário
                saveFileDialog.Filter = "Arquivo Excel (*.xlsx)|*.xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Dados");

                        // Escrever os cabeçalhos das colunas
                        for (int col = 0; col < dataGridView1.Columns.Count; col++)
                        {
                            worksheet.Cell(1, col + 1).Value = dataGridView1.Columns[col].HeaderText;
                        }

                        // Escrever as linhas de dados
                        for (int row = 0; row < dataGridView1.Rows.Count; row++)
                        {
                            for (int col = 0; col < dataGridView1.Columns.Count; col++)
                            {
                                var cellValue = dataGridView1.Rows[row].Cells[col].Value;
                                if (cellValue != null)
                                {
                                    worksheet.Cell(row + 2, col + 1).Value = cellValue.ToString();
                                }
                            }
                        }

                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Dados exportados para Excel com sucesso!");
                    }
                }
            }
        }
    }

}
