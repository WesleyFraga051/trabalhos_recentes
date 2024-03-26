///////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
/////////////Systema dessenvolvido Por Wesley fraga\\\\\\\\\\\\\\\\
///////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Comissao_liquidacao
{
    public partial class Form1 : Form
    {
        string conStr = "User Id=rmsprod;Password=rmsprod;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.1.14)(PORT=1521)))(CONNECT_DATA=(SID=sandbox)))";

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true; // Habilitar a captura de teclas no formulário
            this.KeyDown += Form1_KeyDown; // Associar o evento KeyDown ao manipulador de evento
            checkBox1.CheckedChanged += CheckBox1_CheckedChanged;
            checkBox2.CheckedChanged += CheckBox2_CheckedChanged;
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)
            {
                BtnExec_Click(sender, e);//chama o botao executar
            }
            else if (e.KeyCode == Keys.F3)
            {
                Btnsair_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F10)
            {
                Btnexcel_Click(sender, e);
            }
        }

        private void Variavel(out string dataformat1, out string dataformat2)
        {
            //Data movimento 

            dataformat1 = "1" + dateTimePicker1.Value.ToString("yyMMdd");


            dataformat2 = "1" + dateTimePicker2.Value.ToString("yyMMdd");

        }
        private void BtnExec_Click(object sender, EventArgs e)
        {
            consulta();
        }
        private void Btnsair_Click(object sender, EventArgs e)
        {
            this.Close();// Fechar o Formulario
        }
        private void Btnexcel_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void dateTimePicker1_ValueChanged_1(object sender, EventArgs e)
        {

        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

       /// private string emissao;
        ///private string dtpaga;
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            { MessageBox.Show("Data de emissão selecionada aperte f4 para executar a consulta."); }
            else
            {
                // A caixa de seleção foi desmarcada
                MessageBox.Show("A caixa de seleção foi desmarcada.");
            }
        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            // A caixa de seleção foi marcada
            { MessageBox.Show("Data de pagamento selecionada aperte f4 para executar a consulta."); }
            else
            {
                // A caixa de seleção foi desmarcada
                MessageBox.Show("A caixa de seleção foi desmarcada.");
            }

        }

        private void consulta()
        {
            Variavel(out string dataformat1, out string dataformat2);
            if (!checkBox1.Checked && !checkBox2.Checked)
            {
                MessageBox.Show("Por favor, selecione pelo menos um CheckBox antes de executar a consulta.");
                return; // Retorna para evitar a execução da consulta
            }

            string sqlQuery = "";
            if (checkBox1.Checked)
            {

                // A caixa de seleção foi marcada
                
                sqlQuery = "select distinct" +
                   "\r\nesch_nro_nota as Numero_nfe" +
                   "\r\n,Esch_Agenda As AGENDA" +
                   "\r\n,rms7TO_DATE(esch_data) as Dia_emissao" +
                   "\r\n,Eschlj_Codigo||'-'||Dac(Eschlj_Codigo) as filial_venda" +
                   "\r\n,escl_codigo||'-'||dac(escl_codigo) As Cliente_codigos" +
                   "\r\n,a.tip_razao_social as nome_cliente" +
                   "\r\n,entsai_vendedor||'-'||dac(entsai_vendedor) as vendedor" +
                   "\r\n,c.tip_nome_fantasia as nome_Vendedor" +
                   "\r\n,fis_val_cont_ff_1 As Valor_total_nfe" +
                   "\r\n,entsai_cond_pgto as Condicao_de_pagamento" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 1 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_1" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 1 THEN dup_valor ELSE NULL END) AS Valor_parcela_1" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 1 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_1" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 2 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_2" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 2 THEN dup_valor ELSE NULL END) AS Valor_parcela_2" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 2 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_2" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 3 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_3" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 3 THEN dup_valor ELSE NULL END) AS Valor_parcela_3" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 3 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_3" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 4 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_4" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 4 THEN dup_valor ELSE NULL END) AS Valor_parcela_4" +
                   "\r\n,  MAX(CASE WHEN dup_desd = 4 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_4" +
                   "\r\nfrom aa1rtitu  , ag1fensa,aa2ctipo a , aa2ctipo c, AA1CFISC f" +
                   "\r\n where " +
                   "     esch_data between " + dataformat1 + " and " + dataformat2 + " " +
                   "\r\n--and dup_titulo = 240301 " +
                   "\r\n and dup_agenda in (211,244)" +
                   "\r\n and Eschlj_Codigo = dup_cod_fil" +
                   "\r\n and escl_codigo = dup_cod_cli" +
                   "\r\n and esch_nro_nota = dup_titulo" +
                   "\r\n and esch_agenda = dup_agenda" +
                   "\r\n and a.tip_codigo = escl_codigo" +
                   "\r\n and c.tip_codigo = entsai_vendedor" +
                   "\r\n and f.fis_oper = esch_agenda" +
                   "\r\n and f.fis_nro_nota = esch_nro_nota" +
                   "\r\n and f.fis_dta_agenda = esch_data" +
                   "\r\n GROUP BY" +
                   "\r\n  esch_nro_nota," +
                   "\r\n  Esch_Agenda," +
                   "\r\n  rms7TO_DATE(esch_data)," +
                   "\r\n  Eschlj_Codigo || '-' || Dac(Eschlj_Codigo)," +
                   "\r\n  escl_codigo || '-' || dac(escl_codigo)," +
                   "\r\n  a.tip_razao_social," +
                   "\r\n  entsai_vendedor || '-' || dac(entsai_vendedor)," +
                   "\r\n  c.tip_nome_fantasia," +
                   "\r\n  fis_val_cont_ff_1," +
                   "\r\n  entsai_cond_pgto";

            }
            else if(checkBox2.Checked)
            {
                sqlQuery = "select distinct" +
                  "\r\nesch_nro_nota as Numero_nfe" +
                  "\r\n,Esch_Agenda As AGENDA" +
                  "\r\n,rms7TO_DATE(esch_data) as Dia_emissao" +
                  "\r\n,Eschlj_Codigo||'-'||Dac(Eschlj_Codigo) as filial_venda" +
                  "\r\n,escl_codigo||'-'||dac(escl_codigo) As Cliente_codigos" +
                  "\r\n,a.tip_razao_social as nome_cliente" +
                  "\r\n,entsai_vendedor||'-'||dac(entsai_vendedor) as vendedor" +
                  "\r\n,c.tip_nome_fantasia as nome_Vendedor" +
                  "\r\n,fis_val_cont_ff_1 As Valor_total_nfe" +
                  "\r\n,entsai_cond_pgto as Condicao_de_pagamento" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 1 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_1" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 1 THEN dup_valor ELSE NULL END) AS Valor_parcela_1" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 1 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_1" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 2 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_2" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 2 THEN dup_valor ELSE NULL END) AS Valor_parcela_2" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 2 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_2" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 3 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_3" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 3 THEN dup_valor ELSE NULL END) AS Valor_parcela_3" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 3 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_3" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 4 THEN rms7TO_DATE(DUP_VENC) ELSE NULL END) AS parcela_4" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 4 THEN dup_valor ELSE NULL END) AS Valor_parcela_4" +
                  "\r\n,  MAX(CASE WHEN dup_desd = 4 THEN rms7TO_DATE(dup_dt_pag) ELSE NULL END)  as dia_pag_4" +
                  "\r\nfrom aa1rtitu  , ag1fensa,aa2ctipo a , aa2ctipo c, AA1CFISC f" +
                  "\r\n where " +
                  "     dup_dt_pag between " + dataformat1 + " and " + dataformat2 + " " +
                  "\r\n--and dup_titulo = 240301 " +
                  "\r\n and dup_agenda in (211,244)" +
                  "\r\n and Eschlj_Codigo = dup_cod_fil" +
                  "\r\n and escl_codigo = dup_cod_cli" +
                  "\r\n and esch_nro_nota = dup_titulo" +
                  "\r\n and esch_agenda = dup_agenda" +
                  "\r\n and a.tip_codigo = escl_codigo" +
                  "\r\n and c.tip_codigo = entsai_vendedor" +
                  "\r\n and f.fis_oper = esch_agenda" +
                  "\r\n and f.fis_nro_nota = esch_nro_nota" +
                  "\r\n and f.fis_dta_agenda = esch_data" +
                  "\r\n GROUP BY" +
                  "\r\n  esch_nro_nota," +
                  "\r\n  Esch_Agenda," +
                  "\r\n  rms7TO_DATE(esch_data)," +
                  "\r\n  Eschlj_Codigo || '-' || Dac(Eschlj_Codigo)," +
                  "\r\n  escl_codigo || '-' || dac(escl_codigo)," +
                  "\r\n  a.tip_razao_social," +
                  "\r\n  entsai_vendedor || '-' || dac(entsai_vendedor)," +
                  "\r\n  c.tip_nome_fantasia," +
                  "\r\n  fis_val_cont_ff_1," +
                  "\r\n  entsai_cond_pgto";
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
        private void ExportToExcel()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Não há dados para exportar.");
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.FileName = $"comisaoporliquidação {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx";
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
                                    var cell = worksheet.Cell(row + 2, col + 1);

                                    // Verificar se é uma célula com uma data
                                    if (cellValue is DateTime)
                                    {
                                        cell.Style.NumberFormat.Format = "dd-MM-yyyy";
                                        cell.Value = ((DateTime)cellValue).ToString("dd-MM-yyyy");
                                    }
                                    else
                                    {
                                        cell.Value = cellValue.ToString();
                                    }
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




