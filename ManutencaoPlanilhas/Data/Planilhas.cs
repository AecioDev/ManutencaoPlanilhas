using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ManutencaoPlanilhas.Data
{
    public class Planilhas
    {        
        public bool GeraPlanilhaDeSocio(string caminhoPadrao, int empresaId, string novoNome)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            string PatchPlanilha = caminhoPadrao + "\\" + novoNome + ".xlsx";
            bool continua = false;

            if (!File.Exists(PatchPlanilha))
            {
                var temPlanilha = sheet.RetrieveExcelFromDatabase(empresaId, "S", PatchPlanilha);

                if (temPlanilha > 0)
                {
                    if (temPlanilha == 1)
                    {
                        OpenFileDialog buscaSheet = new OpenFileDialog();

                        // Define o filtro para mostrar apenas arquivos do Excel
                        buscaSheet.Filter = "Arquivos do Excel|*.xlsx";

                        if (buscaSheet.ShowDialog() == DialogResult.OK)
                        {
                            // Obtém o caminho do arquivo selecionado
                            string sheetPath = buscaSheet.FileName;

                            sheet.SaveExcelToDatabase(sheetPath, "S", empresaId);
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

        public bool GeraPlanilhaDeAcerto(string caminhoPadrao, int empresaId, string novoNome)
        {
            string mesAtual, patchNovaPlanilha, mesNumChar;
            bool continua = false;

            for (int i = 1; i <= 12; i++)
            {
                mesNumChar = i > 9 ? i.ToString() + " - " : "0" + i.ToString() + " - ";
                mesAtual = ConverteMes(i);
                //tb_MsgInfo.Text = "Gerando nova Planilha de " + novoNome + " em " + mesAtual + "! Aguarde...";

                string PatchPlanilha = caminhoPadrao + "\\" + mesNumChar + mesAtual;

                if (!Directory.Exists(PatchPlanilha))
                {
                    Directory.CreateDirectory(PatchPlanilha);
                }

                patchNovaPlanilha = PatchPlanilha + "\\" + novoNome + ".xlsx";

                if (!File.Exists(patchNovaPlanilha))
                {
                    PlanilhasBD sheet = new PlanilhasBD();

                    var temPlanilha = sheet.RetrieveExcelFromDatabase(empresaId, "A", patchNovaPlanilha);

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

                                sheet.SaveExcelToDatabase(sheetPath, "A", empresaId);
                            }
                        }
                    }
                    else
                    {
                        continua = true;
                        if (!AlteraNomeDoSocio(patchNovaPlanilha, novoNome, mesAtual, "A", 0))
                        {
                            continua = false;
                            break;
                        }
                    }
                }

                Thread.Sleep(500);
            }

            //tb_MsgInfo.Text = continua ? "Planilhas geradas com Sucesso!" : "Ocorreram erros na Geração das Planilhas";

            return continua;
        }

        public bool AlteraNomeDoSocio(string filePath, string nome, string mes, string tipo, int ano)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            bool iniciouExcel = false;

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
                worksheet = workbook.Sheets[1]; // Abre a aba Janeiro
                iniciouExcel = true;

                if (tipo == "S")
                {
                    // Altera o NOME na Planilha de SÓCIOS
                    worksheet.Cells[2, 2] = nome;

                    if(ano > 0) 
                        worksheet.Cells[2, 10] = ano;

                }
                else
                {
                    // Altera o NOME na Planilha de ACERTO
                    worksheet.Cells[3, 2] = nome;

                    if(!string.IsNullOrWhiteSpace(mes))
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

                if (iniciouExcel)
                {
                    // Libera os objetos COM
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
            }
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
    }
}
