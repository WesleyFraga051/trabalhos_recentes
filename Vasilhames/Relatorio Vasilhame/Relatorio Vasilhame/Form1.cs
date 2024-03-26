using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Text;
using ClosedXML.Excel;

namespace Relatorio_Vasilhame
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
            string sqlQuery = "SELECT " +
                "\r\n    G.GIT_COD_ITEM || '-' || DAC(G.GIT_COD_ITEM) AS CODIGO_ITEM," +
                "\r\n    G.GIT_DESCRICAO AS DESCRIÇÃO," +
                "\r\n    G.GIT_COD_VAS ||'-'|| DAC(G.GIT_COD_VAS) AS CODIGO_VASILHAME," +
                "\r\n    B.GIT_DESCRICAO AS DESCRICAO_VASILHAME," +
                "\r\n    B.GIT_PRC_VEN_1 AS PRECO_VASILHAME " +
                "\r\n    FROM     AA3CITEM G" +
                "\r\n    INNER JOIN AA3CITEM B ON G.GIT_COD_VAS = B.GIT_COD_ITEM" +
                "\r\n    WHERE  G.GIT_COD_VAS <> 0" +
                "\r\n    ORDER BY G.GIT_DESCRICAO ASC";

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
            this.Close();//fechar formulario
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
                saveFileDialog.FileName = "Relatorio Vasilhames.txt";
                saveFileDialog.Filter = "Arquivo Texto (*.txt)|*.txt";
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
                saveFileDialog.FileName = "Relatorio Vasilhames.xlsx"; // Altere o nome conforme necessário
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
