using ManutencaoPlanilhas.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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

            if (!dbOk)
            {
                this.Close();
                return;
            }

            Rb_Socios.Checked = true;
            tb_Ano.Text = DateTime.Now.Year.ToString();
            CarregarComboBoxEmpresas();
        }

        private int ObterEmpresaIdSelecionada()
        {
            if (cb_Empresa.SelectedValue != null && int.TryParse(cb_Empresa.SelectedValue.ToString(), out int id))
            {
                return id;
            }
            return 0;
        }

        private bool TentarObterAnoInformado(out int ano)
        {
            if (string.IsNullOrWhiteSpace(tb_Ano.Text))
            {
                ano = DateTime.Now.Year;
                return true;
            }

            if (int.TryParse(tb_Ano.Text.Trim(), out ano) && ano >= 2000 && ano <= 2100)
            {
                return true;
            }

            MessageBox.Show("Informe um Ano válido (ex: 2026)!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }

        private void cb_Empresa_SelectedIndexChanged(object sender, EventArgs e)
        {
            int empresaId = ObterEmpresaIdSelecionada();
            DefinePatchDefault(empresaId, Rb_Socios.Checked ? "S" : "A");
        }

        private void rb_Socios_CheckedChanged(object sender, EventArgs e)
        {
            if (Rb_Socios.Checked)
                DefinePatchDefault(ObterEmpresaIdSelecionada(), "S");
        }

        private void rb_Acerto_CheckedChanged(object sender, EventArgs e)
        {
            if (Rb_Acerto.Checked)
                DefinePatchDefault(ObterEmpresaIdSelecionada(), "A");
        }

        private void bt_Pasta_Click(object sender, EventArgs e)
        {
            bool selecionou = false;
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                if (!string.IsNullOrWhiteSpace(tb_PastaRaiz.Text) && Directory.Exists(tb_PastaRaiz.Text))
                {
                    folderBrowserDialog.SelectedPath = tb_PastaRaiz.Text;
                }

                DialogResult result = folderBrowserDialog.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    tb_PastaRaiz.Text = folderBrowserDialog.SelectedPath;
                    selecionou = !string.IsNullOrEmpty(tb_PastaRaiz.Text);
                }
                else
                {
                    MessageBox.Show("Nenhum diretório selecionado!", "Atenção",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            if (selecionou && MessageBox.Show("Deseja salvar este caminho como padrão da empresa?", "Atenção!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int empresaId = ObterEmpresaIdSelecionada();
                if (empresaId == 0)
                {
                    MessageBox.Show("Selecione uma Empresa antes de salvar o caminho!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (Directory.Exists(tb_PastaRaiz.Text))
                {
                    PlanilhasBD config = new PlanilhasBD();
                    var keyConfig = Rb_Socios.Checked ? $"DefaultPatchSocio_{empresaId}" : $"DefaultPatchAcerto_{empresaId}";
                    config.SaveConfig(keyConfig, tb_PastaRaiz.Text);
                }
                else
                {
                    MessageBox.Show("A Pasta Principal informada não existe!\n\nSelecione um local válido!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Bt_CadastrarPlanilha_Click(object sender, EventArgs e)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            int empresaId = ObterEmpresaIdSelecionada();

            if (empresaId > 0)
            {
                CadastraNovaPlanilha(empresaId);
            }
            else if (!string.IsNullOrWhiteSpace(cb_Empresa.Text) && cb_Empresa.Text != "Selecione uma Empresa")
            {
                string nomeEmpresa = cb_Empresa.Text.Trim();
                var emp = sheet.CreateEmpresa(nomeEmpresa);
                if (emp)
                {
                    tb_MsgInfo.Text = "Empresa " + nomeEmpresa.ToUpper() + " cadastrada com sucesso!";
                    MessageBox.Show(tb_MsgInfo.Text, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CarregarComboBoxEmpresas();
                }
            }
            else
            {
                MessageBox.Show("Selecione uma Empresa ou informe o Nome para cadastrar!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Bt_GerarPlanilha_Click(object sender, EventArgs e)
        {
            int empresaId = ObterEmpresaIdSelecionada();
            if (empresaId <= 0)
            {
                MessageBox.Show("Selecione uma Empresa!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(tb_PastaRaiz.Text))
            {
                MessageBox.Show("Favor informar a Pasta Principal onde estão as planilhas!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(tb_PastaRaiz.Text))
            {
                MessageBox.Show("A Pasta Principal informada não existe!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!TentarObterAnoInformado(out int anoPlanilha))
            {
                return;
            }

            var nomesInformados = tb_Nome.Text
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(n => n.Trim())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            if (nomesInformados.Count == 0)
            {
                MessageBox.Show("Informe ao menos um nome separado por vírgula!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tb_Nome.Focus();
                return;
            }

            Planilhas plan = new Planilhas();
            int geradas = 0;
            int naoGeradas = 0;

            Bt_GerarPlanilha.Enabled = false;
            try
            {
                foreach (var nome in nomesInformados)
                {
                    tb_MsgInfo.Text = "Gerando planilha de " + nome + "... Por favor aguarde.";
                    Application.DoEvents();

                    PatchPlanilha = Path.Combine(tb_PastaRaiz.Text, nome + ".xlsx");
                    bool ok = false;

                    if (Rb_Socios.Checked)
                    {
                        if (plan.GeraPlanilhaDeSocio(tb_PastaRaiz.Text, empresaId, nome))
                        {
                            ok = plan.AlteraNomeDoSocio(PatchPlanilha, nome, "", "S", anoPlanilha);
                        }
                    }
                    else if (Rb_Acerto.Checked)
                    {
                        ok = plan.GeraPlanilhaDeAcerto(tb_PastaRaiz.Text, empresaId, nome);
                    }

                    if (ok)
                        geradas++;
                    else
                        naoGeradas++;
                }

                if (geradas > 0 && naoGeradas == 0)
                {
                    tb_MsgInfo.Text = $"{geradas} planilha(s) gerada(s) com sucesso!";
                    MessageBox.Show(tb_MsgInfo.Text, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tb_Nome.Text = "";
                }
                else if (geradas > 0)
                {
                    tb_MsgInfo.Text = $"{geradas} gerada(s) e {naoGeradas} ignorada(s) (já existentes ou com erro).";
                    MessageBox.Show(tb_MsgInfo.Text, "Concluído com avisos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    tb_MsgInfo.Text = "Nenhuma planilha foi gerada.";
                    MessageBox.Show("Nenhuma planilha foi gerada (arquivos já existem na pasta ou processo cancelado).", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            finally
            {
                Bt_GerarPlanilha.Enabled = true;
            }
        }

        private void Bt_GeraResumo_Click(object sender, EventArgs e)
        {
            int empresaId = ObterEmpresaIdSelecionada();
            if (empresaId <= 0)
            {
                MessageBox.Show("Selecione uma Empresa!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(tb_PastaRaiz.Text) || !Directory.Exists(tb_PastaRaiz.Text))
            {
                MessageBox.Show("Informe uma Pasta Principal válida antes de gerar o resumo!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!TentarObterAnoInformado(out int anoPlanilha))
            {
                return;
            }

            PlanilhasBD sheet = new PlanilhasBD();
            var tipo = Rb_Socios.Checked ? "S" : "A";
            var resumoOk = sheet.GeraPlanilhaResumo(tb_PastaRaiz.Text, empresaId, anoPlanilha, tipo);

            if (resumoOk)
            {
                MessageBox.Show("Planilha de Resumo criada com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CadastraNovaPlanilha(int empresaId)
        {
            PlanilhasBD sheet = new PlanilhasBD();

            try
            {
                using (OpenFileDialog buscaSheet = new OpenFileDialog())
                {
                    buscaSheet.Filter = "Arquivos do Excel|*.xlsx";

                    if (buscaSheet.ShowDialog() == DialogResult.OK)
                    {
                        string sheetPath = buscaSheet.FileName;
                        var tipo = Rb_Socios.Checked ? "SOCIOS" : "ACERTO";
                        var tipoChar = Rb_Socios.Checked ? "S" : "A";

                        int planilhaGravada = sheet.SaveExcelToDatabase(sheetPath, tipoChar, empresaId);
                        if (planilhaGravada == 0)
                        {
                            tb_MsgInfo.Text = "Planilha Modelo de " + tipo + " atualizada com sucesso!";
                            MessageBox.Show(tb_MsgInfo.Text, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao cadastrar planilha modelo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DefinePatchDefault(int empresa, string tipo)
        {
            tb_PastaRaiz.Text = "";

            if (empresa == 0) return;

            string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            string filePath = Path.Combine(folderPath, "config.txt");

            if (File.Exists(filePath))
            {
                PlanilhasBD config = new PlanilhasBD();
                var keyConfig = tipo == "S" ? $"DefaultPatchSocio_{empresa}" : $"DefaultPatchAcerto_{empresa}";
                tb_PastaRaiz.Text = config.GetConfig(keyConfig) ?? "";
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
            if (string.IsNullOrWhiteSpace(tb_PastaRaiz.Text) || !Directory.Exists(tb_PastaRaiz.Text))
            {
                MessageBox.Show("Favor informar uma Pasta Principal válida das planilhas.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var tipo = Rb_Socios.Checked ? "S" : "A";

            using (AlteraNomes nomes = new AlteraNomes(tb_PastaRaiz.Text, tipo))
            {
                nomes.ShowDialog();
            }
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

        private void ExibirInfoSistema()
        {
            string info = "1 - O Programa utiliza um banco de dados SQLite para armazenar as planilhas de Modelo\n" +
                "\t1.1 - O Banco de dados tem o nome de planilhas.db e fica na pasta Data na raiz do sistema.\n" +
                "2 - Existem 2 tipos de planilhas de Modelo: Sócio (S) e Acerto Semanal (A).\n" +
                "3 - Se o modelo ainda não estiver cadastrado no banco, o sistema solicita a seleção do arquivo e continua automaticamente.\n" +
                "4 - Para cada nome informado, a planilha é gerada na Pasta Principal e tem o Nome (e Ano/Mês) atualizado automaticamente.\n" +
                "5 - Exemplo de arquivo gerado: SOCIO.xlsx";

            MessageBox.Show(info, "Informações do Sistema", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ExibirInfoSistema();
        }

        private void lb_Info_DoubleClick(object sender, EventArgs e)
        {
            ExibirInfoSistema();
        }
    }
}
