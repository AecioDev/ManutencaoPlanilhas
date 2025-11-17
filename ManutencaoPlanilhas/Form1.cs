using ManutencaoPlanilhas.Data;
using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;

namespace ManutencaoPlanilhas
{
    public partial class Form1 : Form
    {
        private string PatchPlanilha = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var migrationManager = new PlanilhasBD();
            var dbOk = migrationManager.ExecMigrations();

            if (!dbOk) { this.Close(); }

            Rb_Socios.Checked = true;   
            //tb_Ano.Text = DateTime.Now.Year.ToString();
            CarregarComboBoxEmpresas();
        }

        private void rb_Socios_CheckedChanged(object sender, EventArgs e)
        {
            if(Rb_Socios.Checked)
                DefinePatchDefault("S");
        }

        private void rb_Acerto_CheckedChanged(object sender, EventArgs e)
        {
            if (Rb_Acerto.Checked)
                DefinePatchDefault("A");
        }

        private void bt_Pasta_Click(object sender, EventArgs e)
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            var selecionou = false;
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
            {
                tb_PastaRaiz.Text = folderBrowserDialog.SelectedPath;
                selecionou = !string.IsNullOrEmpty(tb_PastaRaiz.Text);
            }
            else
            {
                MessageBox.Show("Nenhum diretório selecionado!!!", "Atenção!!!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);                
            }

            if (selecionou && MessageBox.Show("Deseja Salvar este Caminho?", "Atenção!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (Directory.Exists(tb_PastaRaiz.Text))
                {
                    PlanilhasBD config = new PlanilhasBD();
                    var keyConfig = Rb_Socios.Checked ? "DefaultPatchSocio" : "DefaultPatchAcerto";
                    config.SaveConfig(keyConfig, tb_PastaRaiz.Text);                    
                } 
                else
                {
                    MessageBox.Show("A Pasta Principal Informada Não Existe!\n\nSelecione um local válido!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

        }

        private void Bt_CadastrarPlanilha_Click(object sender, EventArgs e)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            int empresaId = cb_Empresa.SelectedValue != null ? Convert.ToInt32(cb_Empresa.SelectedValue) : 0;

            //Verifica se a Empresa Existe
            if (empresaId > 0) //Tem Empresa Selecionada
            {
                CadastraNovaPlanilha(empresaId);
            }
            else if (cb_Empresa.Text != "" && empresaId == 0) //Informou o Nome da Empresa e precisa cadastrar
            {
                var emp = sheet.CreateEmpresa(cb_Empresa.Text);
                if (emp)
                {
                    tb_MsgInfo.Text = "Empresa " + cb_Empresa.Text.ToUpper() + ", Cadastrada com Sucesso!";
                    MessageBox.Show(tb_MsgInfo.Text, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CarregarComboBoxEmpresas();
                }
            }
            else
            {
                MessageBox.Show("Selecione uma Empresa ou informe o Nome para Cadastrar!!!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Bt_GerarPlanilha_Click(object sender, EventArgs e)
        {            
            bool planilhaGerada = false;
            int empresaId = 0;

            empresaId = cb_Empresa.SelectedValue != null ? Convert.ToInt32(cb_Empresa.SelectedValue) : 0;
                        
            if (empresaId <= 0) //Verifica se a Empresa foi Selecionada
            {
                MessageBox.Show("Selecione uma Empresa!!!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!string.IsNullOrEmpty(tb_PastaRaiz.Text))
            {
                if (!Directory.Exists(tb_PastaRaiz.Text))
                {
                    MessageBox.Show("A Pasta Principal Informada Não Existe!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

            }
            else
            {
                MessageBox.Show("Favor informar a Pasta Principal onde estão as outras planilhas!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var AnoPlanilha = string.IsNullOrEmpty(tb_Ano.Text) ? DateTime.Now.Year : Convert.ToInt32(tb_Ano.Text);
            var NomesInformados = tb_Nome.Text.Split(',');
            Planilhas plan = new Planilhas();

            foreach (var nome in NomesInformados)
            {
                tb_MsgInfo.Text = "Gerando Planilha de " + nome.Trim() + "! Por Favor Aguarde...";
                PatchPlanilha = tb_PastaRaiz.Text + "\\" + nome.Trim() + ".xlsx";

                if (Rb_Socios.Checked && plan.GeraPlanilhaDeSocio(tb_PastaRaiz.Text, empresaId, nome.Trim()))
                {
                    planilhaGerada = plan.AlteraNomeDoSocio(PatchPlanilha, nome.Trim(), "", "S", AnoPlanilha);

                    tb_MsgInfo.Text = planilhaGerada ? "Planilha gerada com Sucesso!!!" : "";
                }

                if (Rb_Acerto.Checked)
                {
                    planilhaGerada = plan.GeraPlanilhaDeAcerto(tb_PastaRaiz.Text, empresaId, nome.Trim());
                }
            }

            MessageBox.Show("Planilha (as) Geradas com Sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //Executa Planilha de Resumo.


            this.Close();
        }

        private void Bt_GeraResumo_Click(object sender, EventArgs e)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            var empresaId = cb_Empresa.SelectedValue != null ? Convert.ToInt32(cb_Empresa.SelectedValue) : 0;
            var AnoPlanilha = string.IsNullOrEmpty(tb_Ano.Text) ? DateTime.Now.Year : Convert.ToInt32(tb_Ano.Text);

            if (empresaId > 0)
            {
                var tipo = Rb_Socios.Checked ? "S" : "A";
                var resumoOk = sheet.GeraPlanilhaResumo(tb_PastaRaiz.Text, empresaId, AnoPlanilha, tipo);

                if (resumoOk)
                    MessageBox.Show("Planilha de Resumo criada com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Selecione uma Empresa!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CadastraNovaPlanilha(int empresaId)
        {
            int planilhaGravada;
            PlanilhasBD sheet = new PlanilhasBD();
            OpenFileDialog buscaSheet = new OpenFileDialog();

            try
            {
                // Define o filtro para mostrar apenas arquivos de imagem
                buscaSheet.Filter = "Arquivos do Excel|*.xlsx";

                if (buscaSheet.ShowDialog() == DialogResult.OK)
                {
                    // Obtém o caminho do arquivo selecionado
                    string sheetPath = buscaSheet.FileName;

                    var tipo = Rb_Socios.Checked ? "SOCIOS" : "ACERTO";
                    var tipoChar = Rb_Socios.Checked ? "S" : "A";
                                        
                    planilhaGravada = sheet.SaveExcelToDatabase(sheetPath, tipoChar, empresaId);
                    if (planilhaGravada == 0)
                    {
                        tb_MsgInfo.Text = "Planilha Modelo de " + tipo + ", Atualizada com Sucesso!";
                        MessageBox.Show(tb_MsgInfo.Text, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void DefinePatchDefault(string tipo)
        {
            // Define o caminho para a pasta Data na raiz do sistema
            string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");

            // Caminho completo do arquivo de configurações
            string filePath = Path.Combine(folderPath, "config.txt");

            // Verifica se o arquivo  de configurações existe
            if (File.Exists(filePath))
            {
                PlanilhasBD config = new PlanilhasBD();
                var keyConfig = tipo == "S" ? "DefaultPatchSocio" : "DefaultPatchAcerto";
                tb_PastaRaiz.Text = config.GetConfig(keyConfig);
            }

        }
                
        private void CarregarComboBoxEmpresas()
        {
            EmpresaDB dadosEmp = new EmpresaDB();
            List<Empresa> empresas = dadosEmp.ObterEmpresas();
            cb_Empresa.DataSource = empresas;
            cb_Empresa.DisplayMember = "NomeEmpresa";
            cb_Empresa.ValueMember = "Id";

            cb_Empresa.SelectedIndex = -1;
            cb_Empresa.Text = "Selecione uma Empresa";
        }
        private void bt_AlterarNomes_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(tb_PastaRaiz.Text))
            {
                MessageBox.Show("Favor Informar a Pasta Principal das Planilhas", "Atenção", MessageBoxButtons.OK);
                return;
            }

            var tipo = Rb_Socios.Checked ? "S" : "A";

            AlteraNomes nomes = new AlteraNomes(tb_PastaRaiz.Text, tipo);
            nomes.ShowDialog();
        }

        private void lb_Info_MouseEnter(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "Dois Cliques para Abrir as Informações do Sistema.";
        }

        private void lb_Info_MouseLeave(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "";
        }

        private void tb_Nome_MouseEnter(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "Informe nomes separados por uma vírgula ( Fulano, Ciclano )";
        }

        private void tb_Nome_MouseLeave(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "";
        }

        private void Bt_CadastrarPlanilha_MouseEnter(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "Clique para Cadastrar uma Empresa e Atualizar Planilha Modelo";
        }

        private void Bt_CadastrarPlanilha_MouseLeave(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "";
        }

        private void Form1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string info = "1 - O Programa utiliza um banco de dados SQLite para armazenar as planilhas de Modelo\n" +
                "\t1.1 - O Banco de dados tem o nome de planilhas.db e fica na pasta Data que fica na raíz do sistema.\n" +
                "2 - Existem 2 tipo de planilhas de Modelo a de Sócio e a de Acerto Semanal do Jhonny\n" +
                "3 - O Método BuscaPlanilhaDoBD() recebe uma letra A ou S para buscar a planilha de Sócio ou Acerto\n" +
                "4 - Se a planilha  não existir, o modelo deverá ser incluído no sistema. " +
                "Após o modelo ser inserido e gravado, o sistema continuará normalmente.\n" +
                "5 - Quando o método BuscaPlanilhaDoBD() é executado, o sistema gera a planilha de modelo com o NOME do SÓCIO " +
                "na pasta padrão das planilhas informado no campo de Pasta Principal.\n" +
                "6 - Após a planilha ser gerada na pasta ela é Aberta e Alterada a Celula com o Nome do Sócio, " +
                "e Salva na Pasta Principal com o nome do Sócio. Exemplo: SOCIO.xlsx";

            MessageBox.Show(info, "Informações do Sistema!", MessageBoxButtons.OK);
        }
        private void lb_Info_DoubleClick(object sender, EventArgs e)
        {
            string info = "1 - O Programa utiliza um banco de dados SQLite para armazenar as planilhas de Modelo\n" +
                "\t1.1 - O Banco de dados tem o nome de planilhas.db e fica na pasta Data que fica na raíz do sistema.\n" +
                "2 - Existem 2 tipo de planilhas de Modelo a de Sócio e a de Acerto Semanal do Jhonny\n" +
                "3 - O Método BuscaPlanilhaDoBD() recebe uma letra A ou S para buscar a planilha de Sócio ou Acerto\n" +
                "4 - Se a planilha  não existir, o modelo deverá ser incluído no sistema. " +
                "Após o modelo ser inserido e gravado, o sistema continuará normalmente.\n" +
                "5 - Quando o método BuscaPlanilhaDoBD() é executado, o sistema gera a planilha de modelo com o NOME do SÓCIO " +
                "na pasta padrão das planilhas informado no campo de Pasta Principal.\n" +
                "6 - Após a planilha ser gerada na pasta ela é Aberta e Alterada a Celula com o Nome do Sócio, " +
                "e Salva na Pasta Principal com o nome do Sócio. Exemplo: SOCIO.xlsx";

            MessageBox.Show(info, "Informações do Sistema!", MessageBoxButtons.OK);
        }
    }
}
