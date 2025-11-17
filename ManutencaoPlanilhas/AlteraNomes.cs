using ManutencaoPlanilhas.Data;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ManutencaoPlanilhas
{
    public partial class AlteraNomes : Form
    {
        private string CaminhoPlanilhas;
        private string CaminhoPastas;
        private string TipoPlanilha;
        public AlteraNomes(string _caminhoPlanilhas, string _tipoPlanilha)
        {
            InitializeComponent();
            this.CaminhoPlanilhas = _caminhoPlanilhas;
            this.TipoPlanilha = _tipoPlanilha;
            this.CaminhoPastas = CaminhoPlanilhas;
        }

        private void AlteraNomes_Load(object sender, EventArgs e)
        {
            try
            {
                if(TipoPlanilha == "A")
                {
                    CaminhoPlanilhas = CaminhoPlanilhas + @"\01 - JANEIRO";
                }

                if (!Directory.Exists(CaminhoPlanilhas))
                {
                    MessageBox.Show("O diretório especificado não existe.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Obtém os arquivos com extensão .xlsx e .xls
                var arquivosExcel = Directory.GetFiles(CaminhoPlanilhas, "*.xls*")
                                             .Where(arquivo =>
                                                 !Path.GetFileName(arquivo).StartsWith("~") &&
                                                 !Path.GetFileName(arquivo).StartsWith("#_RESUMO"))
                                             .Select(arquivo => new
                                             {
                                                 Nome = Path.GetFileNameWithoutExtension(arquivo),
                                                 Caminho = TipoPlanilha == "S" ? arquivo : Path.GetFileName(arquivo)
                                             })
                                             .ToList();

                // Configura o DataGridView para permitir edição e inclusão de linhas
                gridNomes.AllowUserToAddRows = true;  // Permite adicionar novas linhas
                gridNomes.AllowUserToDeleteRows = true; // Permite excluir linhas
                gridNomes.EditMode = DataGridViewEditMode.EditOnEnter; // Permite edição ao entrar na célula

                // Configuração das colunas manualmente
                gridNomes.Columns.Clear(); // Limpa colunas anteriores
                gridNomes.ColumnCount = 2; // Define duas colunas

                // Configuração da coluna Nome (editável)
                gridNomes.Columns[0].Name = "nome";
                gridNomes.Columns[0].HeaderText = "Nome";
                gridNomes.Columns[0].Width = 200;
                gridNomes.Columns[0].ReadOnly = false; // Permite edição

                // Configuração da coluna Caminho (não editável)
                gridNomes.Columns[1].Name = "caminho";
                gridNomes.Columns[1].HeaderText = "Caminho Original";
                gridNomes.Columns[1].Width = 400;
                gridNomes.Columns[1].ReadOnly = true; // Caminho original não pode ser alterado pelo usuário

                // Adiciona os arquivos na tabela
                foreach (var arquivo in arquivosExcel)
                {
                    gridNomes.Rows.Add(arquivo.Nome, arquivo.Caminho);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao listar arquivos: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void bt_AlterarNomes_Click(object sender, EventArgs e)
        {
            lb_info.Visible = true;
            bt_AlterarNomes.Enabled = false;

            if (TipoPlanilha == "S") //Planilhas de Sócios
            {
                PlanilhasSocios();
            }
            else
            {
                PlanilhasAcertos();
            }

            lb_info.Text = "Feito!!!";

            MessageBox.Show("Nomes Alterados com Sucesso!!!", "Feito!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            this.Close();
        }

        private void PlanilhasSocios()
        {
            Planilhas plan = new Planilhas();

            foreach (DataGridViewRow row in gridNomes.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string nomeDigitado = row.Cells[0].Value.ToString().Trim(); // Novo nome sem extensão
                    string caminhoOriginal = row.Cells[1].Value.ToString().Trim(); // Caminho do arquivo original

                    lb_info.Text = "Alterando " + nomeDigitado + "...";

                    if (!string.IsNullOrWhiteSpace(nomeDigitado) && File.Exists(caminhoOriginal))
                    {
                        string nomeOriginal = Path.GetFileNameWithoutExtension(caminhoOriginal);

                        nomeOriginal = nomeOriginal.ToUpper();
                        nomeDigitado = nomeDigitado.ToUpper();

                        if (nomeDigitado != nomeOriginal)
                        {
                            string pasta = Path.GetDirectoryName(caminhoOriginal); // Obtém a pasta do arquivo original
                            string extensao = Path.GetExtension(caminhoOriginal);  // Obtém a extensão original

                            string novoCaminho = Path.Combine(pasta, nomeDigitado + extensao); // Cria o novo caminho

                            try
                            {
                                if (!File.Exists(novoCaminho)) // Evita sobrescrever arquivos
                                {
                                    File.Move(caminhoOriginal, novoCaminho); // Renomeia o arquivo
                                    row.Cells[1].Value = novoCaminho; // Atualiza o caminho na grid
                                }

                                plan.AlteraNomeDoSocio(novoCaminho, nomeDigitado, "", "S", 0);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Erro ao renomear o arquivo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }

        private void PlanilhasAcertos()
        {
            Planilhas plan = new Planilhas();

            var AnoMov = DateTime.Now.Year;

            // Loop para percorrer os Meses do Ano
            for (int mes = 1; mes <= 12; mes++)
            {
                var nomeMes = new DateTime(AnoMov, mes, 1).ToString("MMMM"); // Nome do mês
                var mesChar = mes < 10 ? "0" + mes.ToString() : mes.ToString();
                var nomePasta = mesChar + " - " + nomeMes.ToUpper(); //01 - JANEIRO 
                var caminhoPlanilhas = CaminhoPastas + @"\" + nomePasta;

                foreach (DataGridViewRow row in gridNomes.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                    {
                        string nomeDigitado = row.Cells[0].Value.ToString().Trim(); // Novo nome sem extensão
                        string caminhoOriginal = caminhoPlanilhas + @"\" + row.Cells[1].Value.ToString().Trim(); // Nome do arquivo original com extensão

                        lb_info.Text = "Alterando " + nomeDigitado + " em " + nomeMes.ToUpper() + "...";

                        if (!string.IsNullOrWhiteSpace(nomeDigitado) && File.Exists(caminhoOriginal))
                        {
                            string nomeOriginal = Path.GetFileNameWithoutExtension(caminhoOriginal);

                            nomeOriginal = nomeOriginal.ToUpper();
                            nomeDigitado = nomeDigitado.ToUpper();

                            if (nomeDigitado != nomeOriginal)
                            {
                                string pasta = Path.GetDirectoryName(caminhoOriginal); // Obtém a pasta do arquivo original
                                string extensao = Path.GetExtension(caminhoOriginal);  // Obtém a extensão original

                                string novoCaminho = Path.Combine(pasta, nomeDigitado + extensao); // Cria o novo caminho

                                try
                                {
                                    if (!File.Exists(novoCaminho)) // Evita sobrescrever arquivos
                                    {
                                        File.Move(caminhoOriginal, novoCaminho); // Renomeia o arquivo
                                        //row.Cells[1].Value = novoCaminho; // Atualiza o caminho na grid
                                    }

                                    plan.AlteraNomeDoSocio(novoCaminho, nomeDigitado, "", "A", 0);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Erro ao renomear o arquivo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                }
            } //Meses do Ano
        }
    }
}

