using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.SqlServer.Server;
using Oracle.ManagedDataAccess.Client;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace Teste
{
    public partial class Form1 : Form
    {
        string conStr = "User Id=rmsprod;Password=rmsprod;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.1.14)(PORT=1521)))(CONNECT_DATA=(SID=sandbox)))";

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true; // Habilitar a captura de teclas no formulário
            this.KeyDown += Form1_KeyDown; // Associar o evento KeyDown ao manipulador de evento
            comboBox1.Items.Add("Incluido");
            comboBox1.Items.Add("Alterado");
            comboBox1.Items.Add("Baixado");
            comboBox1.Items.Add("Cancelado");
            comboBox2.Items.Add("19 CD");
            comboBox2.Items.Add("27 Mercado");
            comboBox2.Items.Add("35 Atacado");
            comboBox2.Items.Add("51 OpenMall");


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
        private void button1_Click(object sender, EventArgs e)
        {
            //Data movimento 
            var datainicio = dateTimePicker1.Value;
            string datainicioFormat = "1" + dateTimePicker1.Value.ToString("yyMMdd");

            var datafim = dateTimePicker2.Value;
            string datafimFormat = "1" + dateTimePicker2.Value.ToString("yyMMdd");

            //Data Recebimento 
            string recin = "1" + dateTimePicker3.Value.ToString("yyMMdd");


            string recfim = "1" + dateTimePicker4.Value.ToString("yyMMdd");
            //MessageBox.Show($"{recin}");

            //Situação Alterado,Incluido,Baixado,Cancelado
            string selectedItem = comboBox1.SelectedItem.ToString();
            int situacao = 0;
            if (selectedItem == "Incluido")
            {
                situacao = 1;
            }
            else if (selectedItem == "Alterado")
            {
                situacao = 2;
            }
            else if (selectedItem == "Baixado")
            {
                situacao = 3;
            }
            else { situacao = 4; }
            //MessageBox.Show($"Valor numérico associado: {situacao}");

            //FILIAIS
            string filial = comboBox2.SelectedItem.ToString();
            int fil = 0;
            if (filial == "19 CD")
            {
                fil = 1;
            }
            else if (filial == "27 Mercado")
            {
                fil = 2;
            }
            else if (filial == "35 Atacado")
            {
                fil = 3;
            }
            else { fil = 5; }



            // codigo fornecedor 
            string codfor = textBox1.Text;
            MessageBox.Show($"{codfor}");

            //usuario
            string usuario = textBox2.Text;


            using (OracleConnection con = new OracleConnection(conStr))
            {
                con.Open();

                string sqlQuery = "SELECT DISTINCT PED.EXT_COD_LOJ_P || DAC(PED.EXT_COD_LOJ_P) || '-' || LOJ.TIP_NOME_FANTASIA AS LOJA," +
                    "\r\n       PED.EXT_COD_FOR_3 || DAC(PED.EXT_COD_FOR_3) || '-' || FORN.TIP_NOME_FANTASIA AS FORNECEDOR," +
                    "\r\n       PED.EXT_NUM_PED_P || '-' ||DAC( PED.EXT_NUM_PED_P) AS        PEDIDOS," +
                    "\r\n       rms7to_date(PED.EXT_DAT_MOV_P) AS DATA_MOV," +
                    "\r\n       rms7to_date(not_dta_agenda) DATA_RECEBIMENTO," +
                    "\r\n       not_nota as nota," +
                    "\r\n       (CASE PED.EXT_TIP_MOV_P" +
                    "\r\n       WHEN 1 THEN 'INCLUIDO'" +
                    "\r\n       WHEN 2 THEN 'ALTERADO'" +
                    "\r\n       WHEN 3 THEN 'BAIXADO'" +
                    "\r\n       WHEN 4 THEN 'CANCELADO'" +
                    "\r\n       ELSE ' '" +
                    "\r\n       END) AS TIPO_MOV," +
                    "\r\n       PED.EXT_COD_PRO_P || '-' || DAC(PED.EXT_COD_PRO_P) AS PRODUTO," +
                    "\r\n       git_DESCRICAO AS DESCRICAO," +
                    "\r\n       SUBSTR(TRIM(TO_CHAR(PED.EXT_HOR_MOV_P,'000000')),1,2) || ':' || SUBSTR(TRIM(TO_CHAR(PED.EXT_HOR_MOV_P,'000000')),3,2)|| ':' || SUBSTR(TRIM(TO_CHAR(PED.EXT_HOR_MOV_P,'000000')),5,2) AS HORA," +
                    "\r\n       PED.EXT_USUARIO_4," +
                    "\r\n       DECODE(NVL(LENGTH(TRIM(TRANSLATE(EXT_DE_PARA, '+-0123456789.', ' '))),0), 0," +
                    "\r\n              DECODE(PED.EXT_TIP_MOV_P, 2, " +
                    "\r\n              DECODE(PED.EXT_COD_CPO_P, 12, " +
                    "\r\n                     ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,1,13))) || ',' || ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,14,2))), " +
                    "\r\n                     ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,1,15)))), ''), '') AS ANTERIOR," +
                    "\r\n       DECODE(NVL(LENGTH(TRIM(TRANSLATE(EXT_DE_PARA, '+-0123456789.', ' '))),0), 0," +
                    "\r\n       DECODE(PED.EXT_TIP_MOV_P, 2," +
                    "\r\n              DECODE(PED.EXT_COD_CPO_P, 12, " +
                    "\r\n                      ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,16,13))) || ',' || ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,29,2)))," +
                    "\r\n                      ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,16,15)))), ''), '') AS ATUAL, " +
                    "\r\n       DECODE(PED.EXT_TIP_MOV_P, 2," +
                    "\r\n       (CASE PED.EXT_COD_CPO_P" +
                    "\r\n        WHEN 1 THEN 'CND PGTO'" +
                    "\r\n        WHEN 2 THEN 'DT INICIAL'" +
                    "\r\n        WHEN 3 THEN 'DT FIM'" +
                    "\r\n        WHEN 4 THEN 'PGTO'" +
                    "\r\n        WHEN 5 THEN 'FRETE'" +
                    "\r\n        WHEN 6 THEN 'TRANSP'" +
                    "\r\n        WHEN 7 THEN 'OBS'" +
                    "\r\n        WHEN 10 THEN 'EMB'" +
                    "\r\n        WHEN 11 THEN 'QTD PED'" +
                    "\r\n        WHEN 12 THEN 'CUSTO FORN'" +
                    "\r\n        WHEN 13 THEN 'DESCONTO'" +
                    "\r\n        WHEN 14 THEN 'IPI'" +
                    "\r\n        WHEN 15 THEN 'IPI VALOR'" +
                    "\r\n        WHEN 18 THEN 'DESP. ACESS.'  " +
                    "\r\n        WHEN 19 THEN 'COMPRADOR'  " +
                    "\r\n        WHEN 20 THEN 'CUSTO TOT.'  " +
                    "\r\n        ELSE 'OUTRO'" +
                    "\r\n        END), " +
                    "\r\n               ' ') AS CAMPO_ALTERADO, " +
                    "\r\n        (SELECT 1 FROM AG1FLPED " +
                    "\r\n          WHERE CLOJ_CARF   = PED.EXT_COD_LOJ_2" +
                    "\r\n            AND CODFOR_CARF = PED.EXT_COD_FOR_3" +
                    "\r\n            AND NROPED_CARF = PED.EXT_NUM_PED_2) AS SITUACAO" +
                    "\r\nFROM AG1EPEDI PED, AA2CTIPO LOJ, AA2CTIPO FORN, AG2CAPNT, AA3CITEM" +
                    "\r\nWHERE 1 = 1\r\n  and PED.EXT_COD_PRO_P = git_cod_item  " +
                    "\r\n  --and PED.EXT_COD_CPO_P is null" +
                    "\r\n  and PED.EXT_COD_PRO_P <> 0" +
                    "\r\n  AND PED.EXT_TIP_MOV_2   IN (1)" +
                    "\r\n  AND PED.EXT_DAT_MOV_2  between " + datainicioFormat + " and " + datafimFormat + " --datamovimentação" +
                    "\r\n  AND not_dta_agenda BETWEEN " + recin + " and " + recfim + " --recebimento" +
                    "\r\n  AND PED.EXT_TIP_MOV_P = 1" +
                    "\r\n  and EXT_COD_FOR_3 = " + codfor + "--fornecedor" +
                    "\r\n  AND PED.EXT_COD_LOJ_2   in (1,2,3)--filial" +
                    "\r\n  AND EXT_USUARIO_4 LIKE '%ASSC%'--usuario" +
                    "\r\n  AND PED.EXT_TIP_MOV_P = 1" +
                    "\r\n  AND LOJ.TIP_CODIGO      = PED.EXT_COD_LOJ_P" +
                    "\r\n  AND FORN.TIP_CODIGO     = PED.EXT_COD_FOR_3" +
                    "\r\n  and not_nped_1 = PED.EXT_NUM_PED_P" +
                    "\r\n  --and PED.EXT_DAT_MOV_P = not_dta_agenda" +
                    "\r\n  AND (NOT EXISTS (SELECT 1 FROM AG1FLPED " +
                    "\r\n                  WHERE CLOJ_CARF   = PED.EXT_COD_LOJ_2" +
                    "\r\n                    AND CODFOR_CARF = PED.EXT_COD_FOR_3" +
                    "\r\n                    AND NROPED_CARF = PED.EXT_NUM_PED_2) OR PED.EXT_COD_PRO_P = 0)" +
                    "\r\n  --ORDER BY EXT_COD_LOJ_P ASC, EXT_NUM_PED_P ASC, EXT_DAT_MOV_P ASC, EXT_TIP_MOV_P ASC, EXT_COD_PRO_P ASC, EXT_COD_CPO_P ASC, EXT_HOR_MOV_P ASC" +
                    "\r\n union all" +
                    "\r\n SELECT PED.EXT_COD_LOJ_P || DAC(PED.EXT_COD_LOJ_P) || '-' || LOJ.TIP_NOME_FANTASIA AS LOJA," +
                    "\r\n       PED.EXT_COD_FOR_3 || DAC(PED.EXT_COD_FOR_3) || '-' || FORN.TIP_NOME_FANTASIA AS FORNECEDOR," +
                    "\r\n       PED.EXT_NUM_PED_P || '-' ||DAC( PED.EXT_NUM_PED_P) AS        PEDIDOS," +
                    "\r\n       rms7to_date(PED.EXT_DAT_MOV_P) AS DATA_MOV," +
                    "\r\n       rms7to_date(not_dta_agenda) DATA_RECEBIMENTO," +
                    "\r\n       not_nota as nota," +
                    "\r\n       (CASE PED.EXT_TIP_MOV_P" +
                    "\r\n       WHEN 1 THEN 'INCLUIDO'" +
                    "\r\n       WHEN 2 THEN 'ALTERADO'" +
                    "\r\n       WHEN 3 THEN 'BAIXADO'" +
                    "\r\n       WHEN 4 THEN 'CANCELADO'" +
                    "\r\n       ELSE ' '" +
                    "\r\n       END) AS TIPO_MOV," +
                    "\r\n       PED.EXT_COD_PRO_P || '-' || DAC(PED.EXT_COD_PRO_P) AS PRODUTO," +
                    "\r\n       git_DESCRICAO AS DESCRICAO," +
                    "\r\n       SUBSTR(TRIM(TO_CHAR(PED.EXT_HOR_MOV_P,'000000')),1,2) || ':' || SUBSTR(TRIM(TO_CHAR(PED.EXT_HOR_MOV_P,'000000')),3,2)|| ':' || SUBSTR(TRIM(TO_CHAR(PED.EXT_HOR_MOV_P,'000000')),5,2) AS HORA," +
                    "\r\n       PED.EXT_USUARIO_4," +
                    "\r\n       DECODE(NVL(LENGTH(TRIM(TRANSLATE(EXT_DE_PARA, '+-0123456789.', ' '))),0), 0," +
                    "\r\n              DECODE(PED.EXT_TIP_MOV_P, 2, " +
                    "\r\n              DECODE(PED.EXT_COD_CPO_P, 12, " +
                    "\r\n                     ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,1,13))) || ',' || ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,14,2))), " +
                    "\r\n                     ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,1,15)))), ''), '') AS ANTERIOR," +
                    "\r\n       DECODE(NVL(LENGTH(TRIM(TRANSLATE(EXT_DE_PARA, '+-0123456789.', ' '))),0), 0," +
                    "\r\n       DECODE(PED.EXT_TIP_MOV_P, 2," +
                    "\r\n              DECODE(PED.EXT_COD_CPO_P, 12, " +
                    "\r\n                      ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,16,13))) || ',' || ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,29,2)))," +
                    "\r\n                      ROUND(TO_NUMBER(SUBSTR(PED.EXT_DE_PARA,16,15)))), ''), '') AS ATUAL, " +
                    "\r\n       DECODE(PED.EXT_TIP_MOV_P, 2," +
                    "\r\n       (CASE PED.EXT_COD_CPO_P" +
                    "\r\n        WHEN 1 THEN 'CND PGTO'" +
                    "\r\n        WHEN 2 THEN 'DT INICIAL'" +
                    "\r\n        WHEN 3 THEN 'DT FIM'" +
                    "\r\n        WHEN 4 THEN 'PGTO'" +
                    "\r\n        WHEN 5 THEN 'FRETE'" +
                    "\r\n        WHEN 6 THEN 'TRANSP'" +
                    "\r\n        WHEN 7 THEN 'OBS'" +
                    "\r\n        WHEN 10 THEN 'EMB'" +
                    "\r\n        WHEN 11 THEN 'QTD PED'" +
                    "\r\n        WHEN 12 THEN 'CUSTO FORN'" +
                    "\r\n        WHEN 13 THEN 'DESCONTO'" +
                    "\r\n        WHEN 14 THEN 'IPI'" +
                    "\r\n        WHEN 15 THEN 'IPI VALOR'" +
                    "\r\n        WHEN 18 THEN 'DESP. ACESS.'  " +
                    "\r\n        WHEN 19 THEN 'COMPRADOR'  " +
                    "\r\n        WHEN 20 THEN 'CUSTO TOT.'  " +
                    "\r\n        ELSE 'OUTRO'" +
                    "\r\n        END), " +
                    "\r\n        ' ') AS CAMPO_ALTERADO, " +
                    "\r\n        (SELECT 1 FROM AG1FLPED " +
                    "\r\n          WHERE CLOJ_CARF   = PED.EXT_COD_LOJ_2" +
                    "\r\n            AND CODFOR_CARF = PED.EXT_COD_FOR_3" +
                    "\r\n            AND NROPED_CARF = PED.EXT_NUM_PED_2) AS SITUACAO" +
                    "\r\nFROM AG1EPEDI PED, AA2CTIPO LOJ, AA2CTIPO FORN, AG2CAPNT, AA3CITEM" +
                    "\r\nWHERE 1 = 1" +
                    "\r\n  and PED.EXT_COD_CPO_P in (11,12,18)  " +
                    "\r\n   and PED.EXT_COD_PRO_P = git_cod_item " +
                    "\r\n  AND PED.EXT_TIP_MOV_2   IN (2)" +
                    "\r\n  AND PED.EXT_DAT_MOV_2  between " + datainicioFormat + " and " + datafimFormat + " --movimentação" +
                    "\r\n  AND not_dta_agenda BETWEEN " + recin + " and " + recfim + " --recebimento" +
                    "\r\n  AND PED.EXT_TIP_MOV_P = 1" +
                    "\r\n  and EXT_COD_FOR_3 = " + codfor + " --fornecedor" +
                    "\r\n  AND PED.EXT_TIP_MOV_P = 1" +
                    "\r\n  AND EXT_USUARIO_4 LIKE ' %ASSC%'--usuaario" +
                    "\r\n  AND PED.EXT_COD_LOJ_2   in (1,2,3)--filial" +
                    "\r\n  AND LOJ.TIP_CODIGO      = PED.EXT_COD_LOJ_P" +
                    "\r\n  AND FORN.TIP_CODIGO     = PED.EXT_COD_FOR_3" +
                    "\r\n  and not_nped_1 = PED.EXT_NUM_PED_P" +
                    "\r\n  and PED.EXT_DAT_MOV_P = not_dta_agenda" +
                    "\r\n  AND (NOT EXISTS (SELECT 1 FROM AG1FLPED " +
                    "\r\n                  WHERE CLOJ_CARF   = PED.EXT_COD_LOJ_2" +
                    "\r\n                    AND CODFOR_CARF = PED.EXT_COD_FOR_3" +
                    "\r\n                    AND NROPED_CARF = PED.EXT_NUM_PED_2) OR PED.EXT_COD_PRO_P = 0)" +
                    "\r\n  --ORDER BY EXT_COD_LOJ_P ASC, EXT_NUM_PED_P ASC, EXT_DAT_MOV_P ASC, EXT_TIP_MOV_P ASC, EXT_COD_PRO_P ASC, EXT_COD_CPO_P ASC, EXT_HOR_MOV_P ASC\r\n  \r\n";

                OracleDataAdapter adapter = new OracleDataAdapter(sqlQuery, con);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                dataGridView1.DataSource = dt; // Configurar o DataGridView para exibir os dados
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();// Fechar o Formulario a
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }


    }
}