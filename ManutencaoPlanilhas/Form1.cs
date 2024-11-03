using ManutencaoPlanilhas.Data;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Threading;

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
            rb_Socios.Checked = true;   
        }

        private void rb_Socios_CheckedChanged(object sender, EventArgs e)
        {
            if(rb_Socios.Checked)
                DefinePatchDefault("S");
        }

        private void rb_Acerto_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_Acerto.Checked)
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
                    var keyConfig = rb_Socios.Checked ? "DefaultPatchSocio" : "DefaultPatchAcerto";
                    config.SaveConfig(keyConfig, tb_PastaRaiz.Text);                    
                } 
                else
                {
                    MessageBox.Show("A Pasta Principal Informada Não Existe!\n\nSelecione um local válido!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

        }

        private void bt_Adicionar_Click(object sender, EventArgs e)
        {
            bool planilhaGerada = false;

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

            var NomesInformados = tb_Nome.Text.Split(',');

            foreach (var nome in NomesInformados)
            {
                tb_MsgInfo.Text = "Gerando Planilha de " + nome.Trim() + "! Por Favor Aguarde...";

                if (rb_Socios.Checked && GeraPlanilhaDeSocio(nome.Trim()))
                {
                    planilhaGerada = AlteraNomeDoSocio(PatchPlanilha, nome.Trim(), "", "S");

                    tb_MsgInfo.Text = planilhaGerada ? "Planilha gerada com Sucesso!!!" : "";                    
                }

                if (rb_Acerto.Checked)
                {
                    planilhaGerada = GeraPlanilhaDeAcerto(nome.Trim());
                }
            }
                        
            MessageBox.Show("Planilha (as) Geradas com Sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();

        }

        private void bt_NovaPlanilha_DoubleClick(object sender, EventArgs e)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            Int16 planilhaGravada;
            OpenFileDialog buscaSheet = new OpenFileDialog();

            // Define o filtro para mostrar apenas arquivos de imagem
            buscaSheet.Filter = "Arquivos do Excel|*.xlsx";

            if (buscaSheet.ShowDialog() == DialogResult.OK)
            {
                // Obtém o caminho do arquivo selecionado
                string sheetPath = buscaSheet.FileName;

                var tipo = rb_Socios.Checked ? "SOCIOS" : "ACERTO";
                var tipoChar = rb_Socios.Checked ? "S" : "A";

                planilhaGravada = sheet.SaveExcelToDatabase(sheetPath, tipoChar);
                if (planilhaGravada == 0)
                {
                    tb_MsgInfo.Text = "Planilha Modelo de "+ tipo + ", Atualizada com Sucesso!";
                    MessageBox.Show(tb_MsgInfo.Text, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);                    
                }
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

        private bool GeraPlanilhaDeSocio(string novoNome)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            PatchPlanilha = tb_PastaRaiz.Text + "\\" + novoNome + ".xlsx";
            bool continua = false;

            if (!File.Exists(PatchPlanilha))
            {
                var temPlanilha = sheet.RetrieveExcelFromDatabase("S", PatchPlanilha);

                if (temPlanilha > 0)
                {
                    if (temPlanilha == 1)
                    {
                        OpenFileDialog buscaSheet = new OpenFileDialog();

                        // Define o filtro para mostrar apenas arquivos de imagem
                        buscaSheet.Filter = "Arquivos do Excel|*.xlsx";

                        if (buscaSheet.ShowDialog() == DialogResult.OK)
                        {
                            // Obtém o caminho do arquivo selecionado
                            string sheetPath = buscaSheet.FileName;

                            sheet.SaveExcelToDatabase(sheetPath, "S");
                        }
                    }
                }
                else
                {
                    continua = true;
                }
            }

            return continua;
        }

        private bool GeraPlanilhaDeAcerto(string novoNome)
        {
            string mesAtual, patchNovaPlanilha, mesNumChar;
            bool continua = false;

            for (int i = 1; i <= 12; i++)
            {
                mesNumChar = i > 9 ? i.ToString() + " - " : "0" + i.ToString() + " - ";
                mesAtual = ConverteMes(i);
                tb_MsgInfo.Text = "Gerando nova Planilha de " + novoNome + " em " + mesAtual + "! Aguarde...";
                
                PatchPlanilha = tb_PastaRaiz.Text + "\\" + mesNumChar + mesAtual;

                if (!Directory.Exists(PatchPlanilha))
                {
                    Directory.CreateDirectory(PatchPlanilha);
                }

                patchNovaPlanilha = PatchPlanilha + "\\" + novoNome + ".xlsx";

                if (!File.Exists(patchNovaPlanilha))
                {
                    PlanilhasBD sheet = new PlanilhasBD();

                    var temPlanilha = sheet.RetrieveExcelFromDatabase("A", patchNovaPlanilha);

                    if (temPlanilha > 0)
                    {
                        if (temPlanilha == 1) //Não Encontrou a planilha de Exemplo no BD então vai pedir para cadastrar.
                        {
                            OpenFileDialog buscaSheet = new OpenFileDialog();

                            // Define o filtro para mostrar apenas arquivos de imagem
                            buscaSheet.Filter = "Arquivos do Excel|*.xlsx";

                            if (buscaSheet.ShowDialog() == DialogResult.OK)
                            {
                                // Obtém o caminho do arquivo selecionado
                                string sheetPath = buscaSheet.FileName;

                                sheet.SaveExcelToDatabase(sheetPath, "A");
                            }
                        }
                    }
                    else
                    {
                        continua = true;
                        if (!AlteraNomeDoSocio(patchNovaPlanilha, novoNome, mesAtual, "A"))
                        {
                            continua = false;
                            break;
                        }
                    }
                }

                Thread.Sleep(500);
            }

            tb_MsgInfo.Text = continua ? "Planilhas geradas com Sucesso!" : "Ocorreram erros na Geração das Planilhas";

            return continua;
        }

        public bool AlteraNomeDoSocio(string filePath, string name, string mes, string tipo)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Verifica se o arquivo Excel existe
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Planilha não encontrada.\n\nFavor Tentar Novamente", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false; // Falha, arquivo não encontrado
                }

                // Inicializa a aplicação Excel
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = workbook.Sheets[1]; // Abre a primeira planilha

                if (tipo == "S")
                {
                    // Altera o NOME na Planilha de SÓCIOS
                    worksheet.Cells[2, 2] = name;
                }
                else
                {
                    // Altera o NOME na Planilha de ACERTO
                    worksheet.Cells[3, 2] = name;
                    worksheet.Cells[7, 3] = mes;
                }

                // Salva a planilha
                workbook.Save();

                //MessageBox.Show("Planilha gerada com Sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true; // Sucesso
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar a planilha: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Fecha a planilha e a aplicação Excel
                if (workbook != null)
                {
                    workbook.Close(false);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                }

                if (!File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                return false; // Falha
            }
            finally
            {
                // Fecha a planilha e a aplicação Excel
                if (workbook != null)
                {
                    workbook.Close(false);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                }

                // Libera os objetos COM
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
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

        private string ConverteMes(int num)
        {
            string MesChar = "";
            switch (num)
            {
                case 1:
                    MesChar = "JANEIRO";
                    break;
                case 2:
                    MesChar = "FEVEREIRO";
                    break;
                case 3:
                    MesChar = "MARÇO";
                    break;
                case 4:
                    MesChar = "ABRIL";
                    break;
                case 5:
                    MesChar = "MAIO";
                    break;
                case 6:
                    MesChar = "JUNHO";
                    break;
                case 7:
                    MesChar = "JULHO";
                    break;
                case 8:
                    MesChar = "AGOSTO";
                    break;
                case 9:
                    MesChar = "SETEMBRO";
                    break;
                case 10:
                    MesChar = "OUTUBRO";
                    break;
                case 11:
                    MesChar = "NOVEMBRO";
                    break;
                case 12:
                    MesChar = "DEZEMBRO";
                    break;
            }

            return MesChar;
        }

        private void tb_Nome_MouseEnter(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "Informe nomes separados por uma vírgula ( Fulano, Ciclano )";
        }

        private void tb_Nome_MouseLeave(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "";
        }

        private void bt_NovaPlanilha_MouseEnter(object sender, EventArgs e)
        {
            tb_MsgInfo.Text = "Dois cliques para Atualizar Planilha Modelo";
        }

        private void bt_NovaPlanilha_MouseLeave(object sender, EventArgs e)
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
