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
using System.Net;
using System.Net.Mail;
using DocumentFormat.OpenXml.Wordprocessing;


namespace Email
{
    public partial class Form1 : Form
    {
        string conStr = "User Id=rmsprod;Password=rmsprod;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.1.14)(PORT=1521)))(CONNECT_DATA=(SID=sandbox)))";

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true; // Habilitar a captura de teclas no formulário
            this.KeyDown += Form1_KeyDown; // Associar o evento KeyDown ao manipulador de evento
            comboBox1.Items.Add("5");
            comboBox1.Items.Add("10");
            comboBox1.Items.Add("15");
        }

        public void Variavel(out string Dias)
        {
            Dias = comboBox1.SelectedItem.ToString();
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
            consultaemail();
            Consulta();
        }

        private void Email(string destino, string conteudo)
        {

            //cria uma mensagem
            MailMessage mail = new MailMessage();

            //define os endereços
            mail.From = new MailAddress("analista.dev@centermastersul.com.br", "Wesley Fraga");
            //destino
            mail.To.Add(destino);

            //define o conteúdo(assunto)
            mail.Subject = "Este é um simples ,muito simples email";
            //texto do email
            mail.Body = conteudo;
            mail.IsBodyHtml = true;


            //envia a mensagem
            SmtpClient smtp = new SmtpClient("zimbra.centermastersul.com.br");
            smtp.Send(mail);
        }


        private void Consulta()
        {
            string Dias;
            Variavel(out Dias);
            //string filial = fil.ToString();
            string sqlQuery = "select capap_compr as comprador," +
                 " CAPAP_NUMERO as Numero_acordo," +
                 " CAPAP_FORNEC as fornecedor," +
                 " SUBSTR(TO_CHAR(D.DUP_VENC,'fm0000000'),6,2)||'/'||SUBSTR(TO_CHAR(D.DUP_VENC,'fm0000000'),4,2)||'/'||SUBSTR(TO_CHAR(D.DUP_VENC,'fm0000000'),2,2) as vencimento," +
                 " D.DUP_VALOR AS VALOR_DUPLICATA," +
                 //" ((SELECT SUBSTR(TO_CHAR(SYSDATE, 'YY'), 3, 1) || '1' || SUBSTR(TO_CHAR(SYSDATE, 'YYMMDD'), 1)FROM DUAL) -D.DUP_VENC) *0.33 / 100  AS JUROS," +
                // " round(D.DUP_VALOR * ((SELECT SUBSTR(TO_CHAR(SYSDATE, 'YY'), 3, 1) || '1' || SUBSTR(TO_CHAR(SYSDATE, 'YYMMDD'), 1)FROM DUAL) -D.DUP_VENC) *0.33 / 100 + D.DUP_VALOR,2) AS VALOR_COM_JUROS," +
                 " CAPAP_AGENDA AS AGENDA," +
                 " D.DUP_VENC - (SELECT SUBSTR(TO_CHAR(SYSDATE, 'YY'), 3, 1) || '1' || SUBSTR(TO_CHAR(SYSDATE, 'YYMMDD'), 1)FROM DUAL)  as dias_a_vencer" +
                 " FROM AG1CAPAP, AA1RTITU d " +
                 " where" +
                 " CAPAP_NUMERO = DUP_TITULO_ALT" +
                 " and CAPAP_AGENDA = DUP_AGENDA" +
            " AND CAPAP_FLAG_ATU          in(1)" +
                 " AND DUP_VENC  between(SELECT SUBSTR(TO_CHAR(SYSDATE, 'YY'), 3, 1) || '1' || SUBSTR(TO_CHAR(SYSDATE , 'YYMMDD'), 1)FROM DUAL)" +
                 " and(SELECT SUBSTR(TO_CHAR(SYSDATE, 'YY'), 3, 1) || '1' || SUBSTR(TO_CHAR(SYSDATE + "+ Dias + ", 'YYMMDD'), 1)FROM DUAL)" +
                 " AND DUP_DT_PAG = 0" +
                 " ORDER BY DIAS_VENCIDOS DESC ";

            using (OracleConnection con = new OracleConnection(conStr))
            {
                con.Open();
                using (OracleCommand command = new OracleCommand(sqlQuery, con))
                {
                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt); // Preencha o DataTable com os resultados da consulta

                        StringBuilder conteudoEmail = new StringBuilder();
                        conteudoEmail.AppendLine("<html><body><table>");

                        // Adicione a linha de cabeçalho da tabela
                        conteudoEmail.AppendLine("<tr>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Comprador</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Número do Acordo</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Fornecedor</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Vencimento</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Valor da Duplicata</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Valor com Juros</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Agenda</th>");
                        conteudoEmail.AppendLine($"<th style=\"padding: 10px;\">Dias Atrasado</th>");
                        conteudoEmail.AppendLine("</tr>");

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string conteudo1 = dt.Rows[i]["comprador"].ToString();
                            string conteudo2 = dt.Rows[i]["Numero_acordo"].ToString();
                            string conteudo3 = dt.Rows[i]["fornecedor"].ToString();
                            string conteudo4 = dt.Rows[i]["vencimento"].ToString();
                            string conteudo7 = dt.Rows[i]["VALOR_DUPLICATA"].ToString();
                            string conteudo8 = dt.Rows[i]["VALOR_COM_JUROS"].ToString();
                            string conteudo9 = dt.Rows[i]["AGENDA"].ToString();
                            string conteudox = dt.Rows[i]["dias_vencidos"].ToString() ;

                            // Adicione cada linha de dados à tabela com alinhamento central
                            conteudoEmail.AppendLine("<tr>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo1}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo2}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo3}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo4}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo7}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo8}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudo9}</td>");
                            conteudoEmail.AppendLine($"<td align=\"center\">{conteudox}</td>");
                            conteudoEmail.AppendLine("</tr>");
                        }

                        conteudoEmail.AppendLine("</table></body></html>");

                        // Agora você pode enviar o email com as informações em formato HTML
                        string destino = consultaemail(); // Chame consultaemail para obter o destino
                        Email(destino, conteudoEmail.ToString());

                        dataGridView1.DataSource = dt; // Configure o DataGridView para exibir os dados


                    }
                }
            }
        }


        private string consultaemail()
        {
            string query = "SELECT * FROM AA1RMSUS WHERE LTRIM(RTRIM(USU_NOME)) LIKE '%TI03%'";

            using (OracleConnection con = new OracleConnection(conStr))
            {
                con.Open();
                using (OracleCommand command = new OracleCommand(query, con))
                {
                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        StringBuilder destinoEmail = new StringBuilder();

                        foreach (DataRow row in dt.Rows)
                        {
                            string email = row["USU_EMAIL"].ToString();
                            destinoEmail.AppendLine(email);
                        }

                        // Retorna o destino como uma string
                        return destinoEmail.ToString();
                    }
                }
            }
        }
    }
}

