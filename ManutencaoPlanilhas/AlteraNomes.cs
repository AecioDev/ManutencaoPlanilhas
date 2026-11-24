using ManutencaoPlanilhas.Data;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ManutencaoPlanilhas
{
    public partial class AlteraNomes : Form
    {
        private static readonly CultureInfo PtBR = new CultureInfo("pt-BR");

        private string CaminhoPlanilhas;
        private string CaminhoPastas;
        private string TipoPlanilha;

        public AlteraNomes(string _caminhoPlanilhas, string _tipoPlanilha)
        {
            InitializeComponent();
            this.CaminhoPlanilhas = _caminhoPlanilhas;
            this.TipoPlanilha = _tipoPlanilha;
            this.CaminhoPastas = _caminhoPlanilhas;
        }

        private void AlteraNomes_Load(object sender, EventArgs e)
        {
            try
            {
                if (TipoPlanilha == "A")
                {
                    CaminhoPlanilhas = Path.Combine(CaminhoPlanilhas, "01 - JANEIRO");
                }

                if (!Directory.Exists(CaminhoPlanilhas))
                {
                    MessageBox.Show("O diretório especificado não existe.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var arquivosExcel = Directory.GetFiles(CaminhoPlanilhas, "*.xls*")
                                             .Where(arquivo =>
                                             {
                                                 string nome = Path.GetFileName(arquivo);
                                                 return !nome.StartsWith("~") &&
                                                        !nome.StartsWith("#_RESUMO", StringComparison.OrdinalIgnoreCase);
                                             })
                                             .Select(arquivo => new
                                             {
                                                 Nome = Path.GetFileNameWithoutExtension(arquivo),
                                                 Caminho = TipoPlanilha == "S" ? arquivo : Path.GetFileName(arquivo)
                                             })
                                             .ToList();

                gridNomes.AllowUserToAddRows = false;
                gridNomes.AllowUserToDeleteRows = false;
                gridNomes.EditMode = DataGridViewEditMode.EditOnEnter;

                gridNomes.Columns.Clear();
                gridNomes.ColumnCount = 2;

                gridNomes.Columns[0].Name = "nome";
                gridNomes.Columns[0].HeaderText = "Nome";
                gridNomes.Columns[0].Width = 200;
                gridNomes.Columns[0].ReadOnly = false;

                gridNomes.Columns[1].Name = "caminho";
                gridNomes.Columns[1].HeaderText = "Caminho Original";
                gridNomes.Columns[1].Width = 400;
                gridNomes.Columns[1].ReadOnly = true;

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

            try
            {
                int alterados = TipoPlanilha == "S" ? PlanilhasSocios() : PlanilhasAcertos();

                lb_info.Text = "Feito!";
                if (alterados > 0)
                {
                    MessageBox.Show($"{alterados} planilha(s) renomeada(s) com sucesso!", "Feito!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Nenhum nome foi alterado.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            finally
            {
                bt_AlterarNomes.Enabled = true;
            }
        }

        private int PlanilhasSocios()
        {
            Planilhas plan = new Planilhas();
            int alterados = 0;

            foreach (DataGridViewRow row in gridNomes.Rows)
            {
                if (row.IsNewRow || row.Cells[0].Value == null || row.Cells[1].Value == null)
                    continue;

                string nomeDigitado = row.Cells[0].Value.ToString().Trim().ToUpper();
                string caminhoOriginal = row.Cells[1].Value.ToString().Trim();

                if (string.IsNullOrWhiteSpace(nomeDigitado) || !File.Exists(caminhoOriginal))
                    continue;

                string nomeOriginal = Path.GetFileNameWithoutExtension(caminhoOriginal).ToUpper();
                if (nomeDigitado == nomeOriginal)
                    continue;

                lb_info.Text = "Alterando " + nomeDigitado + "...";
                Application.DoEvents();

                string pasta = Path.GetDirectoryName(caminhoOriginal);
                string extensao = Path.GetExtension(caminhoOriginal);
                string novoCaminho = Path.Combine(pasta, nomeDigitado + extensao);

                try
                {
                    if (File.Exists(novoCaminho))
                    {
                        MessageBox.Show($"Já existe um arquivo chamado '{Path.GetFileName(novoCaminho)}'. Alteração ignorada.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    File.Move(caminhoOriginal, novoCaminho);
                    row.Cells[1].Value = novoCaminho;

                    if (plan.AlteraNomeDoSocio(novoCaminho, nomeDigitado, "", "S", 0))
                    {
                        alterados++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro ao renomear o arquivo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return alterados;
        }

        private int PlanilhasAcertos()
        {
            Planilhas plan = new Planilhas();
            int anoMov = DateTime.Now.Year;
            int alterados = 0;

            for (int mes = 1; mes <= 12; mes++)
            {
                string nomeMes = new DateTime(anoMov, mes, 1).ToString("MMMM", PtBR);
                string mesChar = mes < 10 ? "0" + mes : mes.ToString();
                string nomePasta = mesChar + " - " + nomeMes.ToUpper(PtBR);
                string caminhoPastaMes = Path.Combine(CaminhoPastas, nomePasta);

                if (!Directory.Exists(caminhoPastaMes))
                    continue;

                foreach (DataGridViewRow row in gridNomes.Rows)
                {
                    if (row.IsNewRow || row.Cells[0].Value == null || row.Cells[1].Value == null)
                        continue;

                    string nomeDigitado = row.Cells[0].Value.ToString().Trim().ToUpper();
                    string caminhoOriginal = Path.Combine(caminhoPastaMes, row.Cells[1].Value.ToString().Trim());

                    if (string.IsNullOrWhiteSpace(nomeDigitado) || !File.Exists(caminhoOriginal))
                        continue;

                    string nomeOriginal = Path.GetFileNameWithoutExtension(caminhoOriginal).ToUpper();
                    if (nomeDigitado == nomeOriginal)
                        continue;

                    lb_info.Text = "Alterando " + nomeDigitado + " em " + nomeMes.ToUpper(PtBR) + "...";
                    Application.DoEvents();

                    string pasta = Path.GetDirectoryName(caminhoOriginal);
                    string extensao = Path.GetExtension(caminhoOriginal);
                    string novoCaminho = Path.Combine(pasta, nomeDigitado + extensao);

                    try
                    {
                        if (File.Exists(novoCaminho))
                        {
                            MessageBox.Show($"Já existe '{Path.GetFileName(novoCaminho)}' em {nomePasta}. Alteração ignorada.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            continue;
                        }

                        File.Move(caminhoOriginal, novoCaminho);
                        if (plan.AlteraNomeDoSocio(novoCaminho, nomeDigitado, "", "A", 0))
                        {
                            alterados++;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao renomear o arquivo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            return alterados;
        }
    }
}
