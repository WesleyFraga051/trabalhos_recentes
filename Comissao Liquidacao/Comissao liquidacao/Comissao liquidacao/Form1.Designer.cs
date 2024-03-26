namespace Comissao_liquidacao
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            BtnExec = new Button();
            dataGridView1 = new DataGridView();
            Btnsair = new Button();
            Btnexcel = new Button();
            dateTimePicker1 = new DateTimePicker();
            dateTimePicker2 = new DateTimePicker();
            label1 = new Label();
            groupBox1 = new GroupBox();
            checkBox2 = new CheckBox();
            checkBox1 = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // BtnExec
            // 
            BtnExec.Location = new Point(447, 12);
            BtnExec.Name = "BtnExec";
            BtnExec.Size = new Size(124, 53);
            BtnExec.TabIndex = 0;
            BtnExec.Text = "F4  Executar";
            BtnExec.UseVisualStyleBackColor = true;
            BtnExec.Click += BtnExec_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(12, 134);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(1195, 511);
            dataGridView1.TabIndex = 1;
            // 
            // Btnsair
            // 
            Btnsair.Location = new Point(317, 12);
            Btnsair.Name = "Btnsair";
            Btnsair.Size = new Size(124, 53);
            Btnsair.TabIndex = 2;
            Btnsair.Text = "F3 Sair";
            Btnsair.UseVisualStyleBackColor = true;
            Btnsair.Click += Btnsair_Click;
            // 
            // Btnexcel
            // 
            Btnexcel.Location = new Point(577, 12);
            Btnexcel.Name = "Btnexcel";
            Btnexcel.Size = new Size(124, 53);
            Btnexcel.TabIndex = 3;
            Btnexcel.Text = "F10 Excel";
            Btnexcel.UseVisualStyleBackColor = true;
            Btnexcel.Click += Btnexcel_Click;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.CalendarFont = new Font("Sans Serif Collection", 9F, FontStyle.Regular, GraphicsUnit.Point);
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.ImeMode = ImeMode.On;
            dateTimePicker1.Location = new Point(6, 18);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(98, 23);
            dateTimePicker1.TabIndex = 3;
            dateTimePicker1.Value = new DateTime(2023, 11, 8, 0, 0, 0, 0);
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.CalendarFont = new Font("Sans Serif Collection", 9F, FontStyle.Regular, GraphicsUnit.Point);
            dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.ImeMode = ImeMode.On;
            dateTimePicker2.Location = new Point(110, 18);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(98, 23);
            dateTimePicker2.TabIndex = 4;
            dateTimePicker2.Value = new DateTime(2023, 11, 8, 0, 0, 0, 0);
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(14, 69);
            label1.Name = "label1";
            label1.Size = new Size(112, 15);
            label1.TabIndex = 5;
            label1.Text = "Periodo de consulta";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(checkBox2);
            groupBox1.Controls.Add(checkBox1);
            groupBox1.Controls.Add(dateTimePicker2);
            groupBox1.Controls.Add(dateTimePicker1);
            groupBox1.Location = new Point(12, 69);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(334, 59);
            groupBox1.TabIndex = 6;
            groupBox1.TabStop = false;
            // 
            // checkBox2
            // 
            checkBox2.AutoSize = true;
            checkBox2.Location = new Point(214, 34);
            checkBox2.Name = "checkBox2";
            checkBox2.Size = new Size(114, 19);
            checkBox2.TabIndex = 6;
            checkBox2.Text = "Data Pagamento";
            checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(214, 18);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(96, 19);
            checkBox1.TabIndex = 5;
            checkBox1.Text = "Data Emissão";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.LightSkyBlue;
            ClientSize = new Size(1089, 553);
            Controls.Add(label1);
            Controls.Add(Btnexcel);
            Controls.Add(Btnsair);
            Controls.Add(dataGridView1);
            Controls.Add(BtnExec);
            Controls.Add(groupBox1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Relatorio comissão notas e vencimento";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button BtnExec;
        private DataGridView dataGridView1;
        private Button Btnsair;
        private Button Btnexcel;
        private DateTimePicker dateTimePicker1;
        private DateTimePicker dateTimePicker2;
        private Label label1;
        private GroupBox groupBox1;
        private CheckBox checkBox2;
        private CheckBox checkBox1;
    }
}