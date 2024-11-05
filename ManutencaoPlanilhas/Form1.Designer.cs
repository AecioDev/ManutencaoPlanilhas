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
            this.rb_Socios = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rb_Acerto = new System.Windows.Forms.RadioButton();
            this.bt_Adicionar = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_Nome = new System.Windows.Forms.TextBox();
            this.lb_Info = new System.Windows.Forms.Label();
            this.tb_MsgInfo = new System.Windows.Forms.TextBox();
            this.bt_NovaPlanilha = new System.Windows.Forms.PictureBox();
            this.bt_resumo = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.tb_Empresa = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bt_NovaPlanilha)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Pasta Principal";
            // 
            // tb_PastaRaiz
            // 
            this.tb_PastaRaiz.Location = new System.Drawing.Point(15, 122);
            this.tb_PastaRaiz.Name = "tb_PastaRaiz";
            this.tb_PastaRaiz.Size = new System.Drawing.Size(326, 20);
            this.tb_PastaRaiz.TabIndex = 2;
            // 
            // bt_Pasta
            // 
            this.bt_Pasta.Location = new System.Drawing.Point(347, 121);
            this.bt_Pasta.Name = "bt_Pasta";
            this.bt_Pasta.Size = new System.Drawing.Size(30, 22);
            this.bt_Pasta.TabIndex = 2;
            this.bt_Pasta.Text = ". . .";
            this.bt_Pasta.UseVisualStyleBackColor = true;
            this.bt_Pasta.Click += new System.EventHandler(this.bt_Pasta_Click);
            // 
            // rb_Socios
            // 
            this.rb_Socios.AutoSize = true;
            this.rb_Socios.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rb_Socios.Location = new System.Drawing.Point(59, 19);
            this.rb_Socios.Name = "rb_Socios";
            this.rb_Socios.Size = new System.Drawing.Size(97, 17);
            this.rb_Socios.TabIndex = 3;
            this.rb_Socios.Text = "Planilha Sócios";
            this.rb_Socios.UseVisualStyleBackColor = true;
            this.rb_Socios.CheckedChanged += new System.EventHandler(this.rb_Socios_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rb_Acerto);
            this.groupBox1.Controls.Add(this.rb_Socios);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(15, 49);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(362, 45);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tipo";
            // 
            // rb_Acerto
            // 
            this.rb_Acerto.AutoSize = true;
            this.rb_Acerto.Checked = true;
            this.rb_Acerto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rb_Acerto.Location = new System.Drawing.Point(203, 19);
            this.rb_Acerto.Name = "rb_Acerto";
            this.rb_Acerto.Size = new System.Drawing.Size(100, 17);
            this.rb_Acerto.TabIndex = 1;
            this.rb_Acerto.TabStop = true;
            this.rb_Acerto.Text = "Acerto Semanal";
            this.rb_Acerto.UseVisualStyleBackColor = true;
            this.rb_Acerto.CheckedChanged += new System.EventHandler(this.rb_Acerto_CheckedChanged);
            // 
            // bt_Adicionar
            // 
            this.bt_Adicionar.Location = new System.Drawing.Point(306, 169);
            this.bt_Adicionar.Name = "bt_Adicionar";
            this.bt_Adicionar.Size = new System.Drawing.Size(71, 22);
            this.bt_Adicionar.TabIndex = 4;
            this.bt_Adicionar.Text = "Executar";
            this.bt_Adicionar.UseVisualStyleBackColor = true;
            this.bt_Adicionar.Click += new System.EventHandler(this.bt_Adicionar_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 155);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Novo Nome";
            // 
            // tb_Nome
            // 
            this.tb_Nome.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.tb_Nome.Location = new System.Drawing.Point(15, 171);
            this.tb_Nome.Name = "tb_Nome";
            this.tb_Nome.Size = new System.Drawing.Size(158, 20);
            this.tb_Nome.TabIndex = 3;
            this.tb_Nome.MouseEnter += new System.EventHandler(this.tb_Nome_MouseEnter);
            this.tb_Nome.MouseLeave += new System.EventHandler(this.tb_Nome_MouseLeave);
            // 
            // lb_Info
            // 
            this.lb_Info.AutoSize = true;
            this.lb_Info.Location = new System.Drawing.Point(360, 9);
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
            this.tb_MsgInfo.Location = new System.Drawing.Point(12, 197);
            this.tb_MsgInfo.Name = "tb_MsgInfo";
            this.tb_MsgInfo.ReadOnly = true;
            this.tb_MsgInfo.Size = new System.Drawing.Size(362, 15);
            this.tb_MsgInfo.TabIndex = 7;
            this.tb_MsgInfo.Text = "Bem Vindo!";
            this.tb_MsgInfo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // bt_NovaPlanilha
            // 
            this.bt_NovaPlanilha.Image = global::ManutencaoPlanilhas.Properties.Resources.arrow_up;
            this.bt_NovaPlanilha.Location = new System.Drawing.Point(330, 5);
            this.bt_NovaPlanilha.Name = "bt_NovaPlanilha";
            this.bt_NovaPlanilha.Size = new System.Drawing.Size(20, 20);
            this.bt_NovaPlanilha.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.bt_NovaPlanilha.TabIndex = 9;
            this.bt_NovaPlanilha.TabStop = false;
            this.bt_NovaPlanilha.DoubleClick += new System.EventHandler(this.bt_NovaPlanilha_DoubleClick);
            this.bt_NovaPlanilha.MouseEnter += new System.EventHandler(this.bt_NovaPlanilha_MouseEnter);
            this.bt_NovaPlanilha.MouseLeave += new System.EventHandler(this.bt_NovaPlanilha_MouseLeave);
            // 
            // bt_resumo
            // 
            this.bt_resumo.Location = new System.Drawing.Point(229, 169);
            this.bt_resumo.Name = "bt_resumo";
            this.bt_resumo.Size = new System.Drawing.Size(71, 22);
            this.bt_resumo.TabIndex = 4;
            this.bt_resumo.Text = "Resumo";
            this.bt_resumo.UseVisualStyleBackColor = true;
            this.bt_resumo.Click += new System.EventHandler(this.bt_resumo_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Empresa";
            // 
            // tb_Empresa
            // 
            this.tb_Empresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.tb_Empresa.Location = new System.Drawing.Point(15, 21);
            this.tb_Empresa.Name = "tb_Empresa";
            this.tb_Empresa.Size = new System.Drawing.Size(215, 20);
            this.tb_Empresa.TabIndex = 3;
            this.tb_Empresa.Text = "AECIO";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(386, 224);
            this.Controls.Add(this.bt_NovaPlanilha);
            this.Controls.Add(this.tb_MsgInfo);
            this.Controls.Add(this.lb_Info);
            this.Controls.Add(this.bt_resumo);
            this.Controls.Add(this.bt_Adicionar);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.bt_Pasta);
            this.Controls.Add(this.tb_Empresa);
            this.Controls.Add(this.tb_Nome);
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
            ((System.ComponentModel.ISupportInitialize)(this.bt_NovaPlanilha)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_PastaRaiz;
        private System.Windows.Forms.Button bt_Pasta;
        private System.Windows.Forms.RadioButton rb_Socios;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rb_Acerto;
        private System.Windows.Forms.Button bt_Adicionar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_Nome;
        private System.Windows.Forms.Label lb_Info;
        private System.Windows.Forms.TextBox tb_MsgInfo;
        private System.Windows.Forms.PictureBox bt_NovaPlanilha;
        private System.Windows.Forms.Button bt_resumo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tb_Empresa;
    }
}

