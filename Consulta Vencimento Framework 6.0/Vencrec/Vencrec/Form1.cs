///////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
/////////////Systema dessenvolvido Por Wesley fraga\\\\\\\\\\\\\\\\
///////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.SqlServer.Server;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Drawing.Diagrams;


namespace Vencrec
{
    public partial class Form1 : Form
    {
        string conStr = "User Id=rmsprod;Password=rmsprod;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.1.14)(PORT=1521)))(CONNECT_DATA=(SID=sandbox)))";

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true; // Habilitar a captura de teclas no formulário
            this.KeyDown += Form1_KeyDown; // Associar o evento KeyDown ao manipulador de evento
            comboBox1.Items.Add("001 LIMPEZA E HIG");
            comboBox1.Items.Add("002 ELETRO");
            comboBox1.Items.Add("003 PERECIVEIS");
            comboBox1.Items.Add("004 MERCEARIA");
            comboBox1.Items.Add("005 CARNES");
            comboBox1.Items.Add("006 BEBIDAS");
            comboBox1.Items.Add("008 USO CONSUMO");
            comboBox1.Items.Add("009 HORTIFRUTI");
            comboBox1.Items.Add("010 BAZAR");
            comboBox1.Items.Add("999 GERENCIAL");

            comboBox2.Items.Add("19 CD");
            comboBox2.Items.Add("27 Mercado");
            comboBox2.Items.Add("35 Atacado");
            comboBox2.Items.Add("51 OpenMall");
        }
        class Pedidos
        {
            private DateTimePicker dateTimePicker1;
            private DateTimePicker dateTimePicker2;
            private ComboBox comboBox1;
            private ComboBox comboBox2;
            private TextBox textBox1;//fornecedor
            private TextBox textBox3;//agenda
        }

        public void Variavel(out string datainicioFormat, out string datafimFormat, out int comprador, out int fil, out string codfor, out string agenda)
        {


            //Data movimento 

            datainicioFormat = "1" + dateTimePicker1.Value.ToString("yyMMdd");


            datafimFormat = "1" + dateTimePicker2.Value.ToString("yyMMdd");


            //Situação Alterado,Incluido,Baixado,Cancelado
            string selectedItem = comboBox1?.SelectedItem?.ToString(); // Usar ?. para tratar possível ComboBox nulo

            comprador = 0;

            if (selectedItem != null)
            {
                if (selectedItem == "001 LIMPEZA E HIG")
                {
                    comprador = 001;
                }
                else if (selectedItem == "002 ELETRO")
                {
                    comprador = 002;
                }
                else if (selectedItem == "003 PERECIVEIS")
                {
                    comprador = 003;
                }
                else if (selectedItem == "004 MERCEARIA")
                {
                    comprador = 004;
                }
                else if (selectedItem == "005 CARNES")
                {
                    comprador = 005;
                }
                else if (selectedItem == "006 BEBIDAS")
                {
                    comprador = 006;
                }
                else if (selectedItem == "008 USO CONSUMO")
                {
                    comprador = 008;
                }
                else if (selectedItem == "009 HORTIFRUTI")
                {
                    comprador = 009;
                }
                else if (selectedItem == "010 BAZAR")
                {
                    comprador = 010;
                }
                else if (selectedItem == "999 GERENCIAL")
                {
                    comprador = 999;
                }
                else
                {
                    // Lógica para tratamento de valor não reconhecido
                }
            }
            else
            {
                // Lógica para tratamento de ComboBox nulo
            }
            //MessageBox.Show($"Valor numérico associado: {comprador}");

            //FILIAIS
            string filial = comboBox2?.SelectedItem?.ToString();

            fil = 0;

            if (filial != null)
            {
                if (filial == "19 CD")
                {
                    fil = 19;
                }
                else if (filial == "27 Mercado")
                {
                    fil = 27;
                }
                else if (filial == "35 Atacado")
                {
                    fil = 35;
                }
                else
                {
                    fil = 51;
                }
            }
            else
            {
            }



            // codigo fornecedor 
            codfor = textBox1.Text;
            // MessageBox.Show($"{ codfor}");

            //usuario
            agenda = textBox3.Text;

        }


        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)
            {
                button1_Click(sender, e); // Chamar o evento do botão executar
            }
            else if (e.KeyCode == Keys.F3)
            {
                button2_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F6)
            {
                button3_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F10)
            {
                button4_Click(sender, e);
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();// Fechar o Formulario
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            Consulta();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportTotxt();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }
        private void dateTimePicker1_ValueChanged_1(object sender, EventArgs e)
        {

        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

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
                saveFileDialog.FileName = $"Vencimento {DateTime.Now.ToString("dd-MM-yyyy")}.txt";
                saveFileDialog.Filter = " (*.txt)|*.txt";
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
                saveFileDialog.FileName = $"Vencimento {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx"; // Altere o nome conforme necessário
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

        private void Consulta()
        {
           

            Variavel(out string datainicioFormat, out string datafimFormat, out int comprador, out int fil, out string codfor, out string agenda);
            //string filial = fil.ToString();
            string sqlQuery = "SELECT " +
                 "CODIGO, " +
                 "DESCRICAO, " +
                 "DATA_AGENDA, " +
                 "LIMITE, " +
                 "FABRICACAO, " +
                 "VALIDADE, " +
                 "AGENDA, " +
                 "NOTA, " +
                 "FORNECEDOR, " +
                 "FILIAL, " +
                 "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
             "FROM (" +
                 "SELECT DISTINCT " +
                     "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                     "GIT_DESCRICAO AS DESCRICAO, " +
                     "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                     "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                     "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                     "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                     "REN_AGE_ITEM AS AGENDA, " +
                     "REN_NOTA AS NOTA, " +
                     "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                     "REN_DESTINO AS FILIAL, " +
                     "REN_QTD_REC, " +
                     "GIT_EMB_FOR " +
                 "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                 "WHERE " +
                     "GIT_COD_ITEM = DET_COD_ITEM " +
                     "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                     "AND REN_DISTRIB = TIP_CODIGO " +
                     "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                     //"AND REN_DESTINO = " + fil + " " +
                     "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
             ") subquery " +
             "GROUP BY " +
                 "CODIGO, " +
                 "DESCRICAO, " +
                 "DATA_AGENDA, " +
                 "LIMITE, " +
                 "FABRICACAO, " +
                 "VALIDADE, " +
                 "AGENDA, " +
                 "NOTA, " +
                 "FORNECEDOR, " +
                 "FILIAL";

            if (comboBox1.Text == "" && textBox3.Text == "" && textBox1.Text == "" && comboBox2.Text != "")
            //filtro apenas filial entrada
            {
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + " " +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text == "" && textBox3.Text != "" && textBox1.Text == "" && comboBox2.Text == "")
            //filtro apenas agenda entrada
            {
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_AGE_ITEM = " + agenda + " " +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text == "" && textBox1.Text == "" && comboBox2.Text == "")
            //filtro apenas comprador 
            {
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND GIT_COMPRADOR = " + comprador + " " +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text != "" && textBox1.Text != "" && comboBox2.Text != "")
            //filtro apenas todos filtros
            {
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + " " +
                    "AND REN_AGE_ITEM = " + agenda + " " +
                    "AND GIT_COMPRADOR = " + comprador + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text == "" && textBox3.Text != "" && textBox1.Text == "" && comboBox2.Text != "")
            {   //filtro apenas filial agenda
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + "" +
                    "AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text == "" && textBox1.Text == "" && comboBox2.Text != "")
            {   //filtro apenas filial comprador
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + "" +
                    "AND  GIT_COMPRADOR= " + comprador + "" +
                    //"AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text != "" && textBox1.Text == "" && comboBox2.Text == "")
            {   //filtro apenas agenda comprador
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    //"AND REN_DESTINO = " + fil + "" +
                    "AND  GIT_COMPRADOR= " + comprador + "" +
                    "AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text == "" && textBox3.Text == "" && textBox1.Text != "" && comboBox2.Text == "")
            {   //filtro apenas codigo fornecedor
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DISTRIB = " + codfor + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text == "" && textBox1.Text != "" && comboBox2.Text == "")
            {   //filtro apenas fornecedor comprador
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    //"AND REN_DESTINO = " + fil + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    "AND  GIT_COMPRADOR= " + comprador + "" +
                    //"AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text == "" && textBox1.Text != "" && comboBox2.Text != "")
            {   //filtro apenas fornecedor,comprador, filial
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    "AND  GIT_COMPRADOR= " + comprador + "" +
                    //"AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text == "" && textBox3.Text == "" && textBox1.Text != "" && comboBox2.Text != "")
            {   //filtro apenas fornecedor, filial
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    //"AND  GIT_COMPRADOR= " + comprador + "" +
                    //"AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text == "" && textBox3.Text != "" && textBox1.Text != "" && comboBox2.Text != "")
            {   //filtro apenas fornecedor, filial agenda
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    //"AND  GIT_COMPRADOR= " + comprador + "" +
                    "AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text == "" && textBox3.Text != "" && textBox1.Text != "" && comboBox2.Text == "")
            {   //filtro apenas fornecedor,  agenda
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    // "AND REN_DESTINO = " + fil + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    //"AND  GIT_COMPRADOR= " + comprador + "" +
                    "AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text != "" && textBox1.Text != "" && comboBox2.Text == "")
            {   //filtro apenas fornecedor,  agenda,comprador
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    // "AND REN_DESTINO = " + fil + "" +
                    "AND REN_DISTRIB = " + codfor + "" +
                    "AND  GIT_COMPRADOR= " + comprador + "" +
                    "AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else if (comboBox1.Text != "" && textBox3.Text != "" && textBox1.Text == "" && comboBox2.Text != "")
            {   //filtro filial agenda comprador 
                sqlQuery = "SELECT " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL, " +
                "SUM(REN_QTD_REC * GIT_EMB_FOR) AS QTD_RECEBIDA " +
            "FROM (" +
                "SELECT DISTINCT " +
                    "GIT_COD_ITEM || '-' || GIT_DIGITO AS CODIGO, " +
                    "GIT_DESCRICAO AS DESCRICAO, " +
                    "RMS7TO_DATE(REN_DTA_AGENDA) AS DATA_AGENDA, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_RECEBIMENTO AS LIMITE, " +
                    "RMSTO_DATE(REN_DaTA_validade) - DET_DIAS_EXPEDICAO AS FABRICACAO, " +
                    "RMSTO_DATE(REN_DaTA_validade) AS VALIDADE, " +
                    "REN_AGE_ITEM AS AGENDA, " +
                    "REN_NOTA AS NOTA, " +
                    "TIP_RAZAO_SOCIAL AS FORNECEDOR, " +
                    "REN_DESTINO AS FILIAL, " +
                    "REN_QTD_REC, " +
                    "GIT_EMB_FOR " +
                "FROM AA3CITEM, AA1DITEM, AG2DETNT, AA2CTIPO " +
                "WHERE " +
                    "GIT_COD_ITEM = DET_COD_ITEM " +
                    "AND REN_DTA_AGENDA BETWEEN " + datainicioFormat + " and " + datafimFormat + " " +
                    "AND REN_DISTRIB = TIP_CODIGO " +
                    "AND DET_CONTROLE_LOTE_SERIE = '4' " +
                    "AND REN_DESTINO = " + fil + "" +
                    //"AND REN_DISTRIB = " + codfor + "" +
                    "AND  GIT_COMPRADOR= " + comprador + "" +
                    "AND REN_AGE_ITEM = " + agenda + "" +
                    "AND GIT_COD_ITEM = SUBSTR(TO_CHAR(REN_ITEM, 'fm000000000000000'), 8, 7) " +
            ") subquery " +
            "GROUP BY " +
                "CODIGO, " +
                "DESCRICAO, " +
                "DATA_AGENDA, " +
                "LIMITE, " +
                "FABRICACAO, " +
                "VALIDADE, " +
                "AGENDA, " +
                "NOTA, " +
                "FORNECEDOR, " +
                "FILIAL";
            }
            else
            {
                MessageBox.Show("TODAS AS FILIAIS, COMPRADORES,AGENDAS E FORNECEDORES  ");
            }

            using (OracleConnection con = new OracleConnection(conStr))
            {
                con.Open();
                using (OracleCommand command = new OracleCommand(sqlQuery, con))
                {
                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt); // Preencha o DataTable com os resultados da consulta

                        dataGridView1.DataSource = dt; // Configure o DataGridView para exibir os dados
                    }
                }
            }


        }

       
    }
}
