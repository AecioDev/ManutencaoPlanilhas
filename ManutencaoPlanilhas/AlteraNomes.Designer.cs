namespace ManutencaoPlanilhas
{
    partial class AlteraNomes
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
            this.gridNomes = new System.Windows.Forms.DataGridView();
            this.nome = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.bt_AlterarNomes = new System.Windows.Forms.Button();
            this.lb_info = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.gridNomes)).BeginInit();
            this.SuspendLayout();
            // 
            // gridNomes
            // 
            this.gridNomes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridNomes.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.nome});
            this.gridNomes.Location = new System.Drawing.Point(21, 73);
            this.gridNomes.Name = "gridNomes";
            this.gridNomes.Size = new System.Drawing.Size(269, 220);
            this.gridNomes.TabIndex = 0;
            // 
            // nome
            // 
            this.nome.HeaderText = "Nome";
            this.nome.Name = "nome";
            this.nome.Width = 200;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(11, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(288, 42);
            this.label1.TabIndex = 1;
            this.label1.Text = "Altere os Nomes na Grid e Clique no \r\nBotão Alterar";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // bt_AlterarNomes
            // 
            this.bt_AlterarNomes.Location = new System.Drawing.Point(107, 299);
            this.bt_AlterarNomes.Name = "bt_AlterarNomes";
            this.bt_AlterarNomes.Size = new System.Drawing.Size(97, 26);
            this.bt_AlterarNomes.TabIndex = 2;
            this.bt_AlterarNomes.Text = "Alterar Nomes";
            this.bt_AlterarNomes.UseVisualStyleBackColor = true;
            this.bt_AlterarNomes.Click += new System.EventHandler(this.bt_AlterarNomes_Click);
            // 
            // lb_info
            // 
            this.lb_info.AutoSize = true;
            this.lb_info.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_info.ForeColor = System.Drawing.Color.DarkRed;
            this.lb_info.Location = new System.Drawing.Point(61, 53);
            this.lb_info.Name = "lb_info";
            this.lb_info.Size = new System.Drawing.Size(188, 17);
            this.lb_info.TabIndex = 3;
            this.lb_info.Text = "Alterando Nomes, Aguarde...";
            this.lb_info.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lb_info.Visible = false;
            // 
            // AlteraNomes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(310, 333);
            this.Controls.Add(this.lb_info);
            this.Controls.Add(this.bt_AlterarNomes);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.gridNomes);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "AlteraNomes";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AlteraNomes";
            this.Load += new System.EventHandler(this.AlteraNomes_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridNomes)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView gridNomes;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bt_AlterarNomes;
        private System.Windows.Forms.DataGridViewTextBoxColumn nome;
        private System.Windows.Forms.Label lb_info;
    }
}