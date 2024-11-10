namespace ManutencaoPlanilhas
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.tb_PastaRaiz = new System.Windows.Forms.TextBox();
            this.bt_Pasta = new System.Windows.Forms.Button();
            this.Rb_Socios = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Rb_Acerto = new System.Windows.Forms.RadioButton();
            this.Bt_GerarPlanilha = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_Nome = new System.Windows.Forms.TextBox();
            this.lb_Info = new System.Windows.Forms.Label();
            this.tb_MsgInfo = new System.Windows.Forms.TextBox();
            this.Bt_GeraResumo = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.tb_Ano = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cb_Empresa = new System.Windows.Forms.ComboBox();
            this.Bt_CadastrarPlanilha = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 118);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Pasta Principal";
            // 
            // tb_PastaRaiz
            // 
            this.tb_PastaRaiz.Location = new System.Drawing.Point(12, 134);
            this.tb_PastaRaiz.Name = "tb_PastaRaiz";
            this.tb_PastaRaiz.Size = new System.Drawing.Size(356, 20);
            this.tb_PastaRaiz.TabIndex = 2;
            // 
            // bt_Pasta
            // 
            this.bt_Pasta.Location = new System.Drawing.Point(374, 133);
            this.bt_Pasta.Name = "bt_Pasta";
            this.bt_Pasta.Size = new System.Drawing.Size(30, 22);
            this.bt_Pasta.TabIndex = 2;
            this.bt_Pasta.Text = ". . .";
            this.bt_Pasta.UseVisualStyleBackColor = true;
            this.bt_Pasta.Click += new System.EventHandler(this.bt_Pasta_Click);
            // 
            // Rb_Socios
            // 
            this.Rb_Socios.AutoSize = true;
            this.Rb_Socios.Checked = true;
            this.Rb_Socios.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Rb_Socios.Location = new System.Drawing.Point(74, 19);
            this.Rb_Socios.Name = "Rb_Socios";
            this.Rb_Socios.Size = new System.Drawing.Size(97, 17);
            this.Rb_Socios.TabIndex = 3;
            this.Rb_Socios.TabStop = true;
            this.Rb_Socios.Text = "Planilha Sócios";
            this.Rb_Socios.UseVisualStyleBackColor = true;
            this.Rb_Socios.CheckedChanged += new System.EventHandler(this.rb_Socios_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Rb_Acerto);
            this.groupBox1.Controls.Add(this.Rb_Socios);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 61);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(392, 45);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tipo";
            // 
            // Rb_Acerto
            // 
            this.Rb_Acerto.AutoSize = true;
            this.Rb_Acerto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Rb_Acerto.Location = new System.Drawing.Point(218, 19);
            this.Rb_Acerto.Name = "Rb_Acerto";
            this.Rb_Acerto.Size = new System.Drawing.Size(100, 17);
            this.Rb_Acerto.TabIndex = 1;
            this.Rb_Acerto.Text = "Acerto Semanal";
            this.Rb_Acerto.UseVisualStyleBackColor = true;
            this.Rb_Acerto.CheckedChanged += new System.EventHandler(this.rb_Acerto_CheckedChanged);
            // 
            // Bt_GerarPlanilha
            // 
            this.Bt_GerarPlanilha.Location = new System.Drawing.Point(128, 214);
            this.Bt_GerarPlanilha.Name = "Bt_GerarPlanilha";
            this.Bt_GerarPlanilha.Size = new System.Drawing.Size(88, 22);
            this.Bt_GerarPlanilha.TabIndex = 4;
            this.Bt_GerarPlanilha.Text = "Gerar Planilhas";
            this.Bt_GerarPlanilha.UseVisualStyleBackColor = true;
            this.Bt_GerarPlanilha.Click += new System.EventHandler(this.Bt_GerarPlanilha_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(9, 167);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(283, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Informe os Nomes abaixo Separados por vírgula.";
            // 
            // tb_Nome
            // 
            this.tb_Nome.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.tb_Nome.Location = new System.Drawing.Point(12, 183);
            this.tb_Nome.Name = "tb_Nome";
            this.tb_Nome.Size = new System.Drawing.Size(392, 20);
            this.tb_Nome.TabIndex = 3;
            this.tb_Nome.MouseEnter += new System.EventHandler(this.tb_Nome_MouseEnter);
            this.tb_Nome.MouseLeave += new System.EventHandler(this.tb_Nome_MouseLeave);
            // 
            // lb_Info
            // 
            this.lb_Info.AutoSize = true;
            this.lb_Info.Location = new System.Drawing.Point(391, 9);
            this.lb_Info.Name = "lb_Info";
            this.lb_Info.Size = new System.Drawing.Size(13, 13);
            this.lb_Info.TabIndex = 6;
            this.lb_Info.Text = "?";
            this.lb_Info.DoubleClick += new System.EventHandler(this.lb_Info_DoubleClick);
            this.lb_Info.MouseEnter += new System.EventHandler(this.lb_Info_MouseEnter);
            this.lb_Info.MouseLeave += new System.EventHandler(this.lb_Info_MouseLeave);
            // 
            // tb_MsgInfo
            // 
            this.tb_MsgInfo.BackColor = System.Drawing.SystemColors.Control;
            this.tb_MsgInfo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tb_MsgInfo.Font = new System.Drawing.Font("Lato", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tb_MsgInfo.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.tb_MsgInfo.Location = new System.Drawing.Point(12, 246);
            this.tb_MsgInfo.Name = "tb_MsgInfo";
            this.tb_MsgInfo.ReadOnly = true;
            this.tb_MsgInfo.Size = new System.Drawing.Size(392, 15);
            this.tb_MsgInfo.TabIndex = 7;
            this.tb_MsgInfo.Text = "Bem Vindo!";
            this.tb_MsgInfo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Bt_GeraResumo
            // 
            this.Bt_GeraResumo.Location = new System.Drawing.Point(222, 214);
            this.Bt_GeraResumo.Name = "Bt_GeraResumo";
            this.Bt_GeraResumo.Size = new System.Drawing.Size(88, 22);
            this.Bt_GeraResumo.TabIndex = 4;
            this.Bt_GeraResumo.Text = "Gerar Resumo";
            this.Bt_GeraResumo.UseVisualStyleBackColor = true;
            this.Bt_GeraResumo.Click += new System.EventHandler(this.Bt_GeraResumo_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(9, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Empresa";
            // 
            // tb_Ano
            // 
            this.tb_Ano.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.tb_Ano.Location = new System.Drawing.Point(232, 32);
            this.tb_Ano.Name = "tb_Ano";
            this.tb_Ano.Size = new System.Drawing.Size(65, 20);
            this.tb_Ano.TabIndex = 3;
            this.tb_Ano.Text = "2025";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(229, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Ano";
            // 
            // cb_Empresa
            // 
            this.cb_Empresa.FormattingEnabled = true;
            this.cb_Empresa.Location = new System.Drawing.Point(12, 32);
            this.cb_Empresa.Name = "cb_Empresa";
            this.cb_Empresa.Size = new System.Drawing.Size(214, 21);
            this.cb_Empresa.TabIndex = 10;
            // 
            // Bt_CadastrarPlanilha
            // 
            this.Bt_CadastrarPlanilha.Location = new System.Drawing.Point(303, 31);
            this.Bt_CadastrarPlanilha.Name = "Bt_CadastrarPlanilha";
            this.Bt_CadastrarPlanilha.Size = new System.Drawing.Size(101, 22);
            this.Bt_CadastrarPlanilha.TabIndex = 11;
            this.Bt_CadastrarPlanilha.Text = "Cadastrar Planilha";
            this.Bt_CadastrarPlanilha.UseVisualStyleBackColor = true;
            this.Bt_CadastrarPlanilha.Click += new System.EventHandler(this.Bt_CadastrarPlanilha_Click);
            this.Bt_CadastrarPlanilha.MouseEnter += new System.EventHandler(this.Bt_CadastrarPlanilha_MouseEnter);
            this.Bt_CadastrarPlanilha.MouseLeave += new System.EventHandler(this.Bt_CadastrarPlanilha_MouseLeave);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(418, 273);
            this.Controls.Add(this.Bt_CadastrarPlanilha);
            this.Controls.Add(this.cb_Empresa);
            this.Controls.Add(this.tb_MsgInfo);
            this.Controls.Add(this.lb_Info);
            this.Controls.Add(this.Bt_GeraResumo);
            this.Controls.Add(this.Bt_GerarPlanilha);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.bt_Pasta);
            this.Controls.Add(this.tb_Ano);
            this.Controls.Add(this.tb_Nome);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tb_PastaRaiz);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Incluir Planilha";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Form1_MouseDoubleClick);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_PastaRaiz;
        private System.Windows.Forms.Button bt_Pasta;
        private System.Windows.Forms.RadioButton Rb_Socios;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton Rb_Acerto;
        private System.Windows.Forms.Button Bt_GerarPlanilha;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_Nome;
        private System.Windows.Forms.Label lb_Info;
        private System.Windows.Forms.Button Bt_GeraResumo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tb_Ano;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cb_Empresa;
        private System.Windows.Forms.Button Bt_CadastrarPlanilha;
        public System.Windows.Forms.TextBox tb_MsgInfo;
    }
}

