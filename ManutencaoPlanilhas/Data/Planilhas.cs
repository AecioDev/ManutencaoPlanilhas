using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace ManutencaoPlanilhas.Data
{
    public class Planilhas
    {
        public bool GeraPlanilhaDeSocio(string caminhoPadrao, int empresaId, string novoNome)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            string patchPlanilha = Path.Combine(caminhoPadrao, novoNome + ".xlsx");

            if (File.Exists(patchPlanilha))
            {
                return false;
            }

            var temPlanilha = sheet.RetrieveExcelFromDatabase(empresaId, "S", patchPlanilha);

            if (temPlanilha == 1)
            {
                if (SolicitarECadastrarModelo(sheet, "S", empresaId))
                {
                    temPlanilha = sheet.RetrieveExcelFromDatabase(empresaId, "S", patchPlanilha);
                }
            }

            return temPlanilha == 0;
        }

        public bool GeraPlanilhaDeAcerto(string caminhoPadrao, int empresaId, string novoNome)
        {
            PlanilhasBD sheet = new PlanilhasBD();
            bool algumaGerada = false;

            for (int i = 1; i <= 12; i++)
            {
                string mesNumChar = i > 9 ? i + " - " : "0" + i + " - ";
                string mesAtual = ConverteMes(i);

                string pastaMes = Path.Combine(caminhoPadrao, mesNumChar + mesAtual);
                if (!Directory.Exists(pastaMes))
                {
                    Directory.CreateDirectory(pastaMes);
                }

                string patchNovaPlanilha = Path.Combine(pastaMes, novoNome + ".xlsx");
                if (File.Exists(patchNovaPlanilha))
                {
                    continue;
                }

                var temPlanilha = sheet.RetrieveExcelFromDatabase(empresaId, "A", patchNovaPlanilha);

                if (temPlanilha == 1)
                {
                    if (!SolicitarECadastrarModelo(sheet, "A", empresaId))
                    {
                        // Usuário cancelou a seleção do modelo; interrompe os meses restantes
                        return algumaGerada;
                    }
                    temPlanilha = sheet.RetrieveExcelFromDatabase(empresaId, "A", patchNovaPlanilha);
                }

                if (temPlanilha == 0)
                {
                    if (!AlteraNomeDoSocio(patchNovaPlanilha, novoNome, mesAtual, "A", 0))
                    {
                        return false;
                    }
                    algumaGerada = true;
                }
                else
                {
                    return false;
                }
            }

            return algumaGerada;
        }

        private bool SolicitarECadastrarModelo(PlanilhasBD sheet, string tipo, int empresaId)
        {
            using (OpenFileDialog buscaSheet = new OpenFileDialog())
            {
                buscaSheet.Filter = "Arquivos do Excel|*.xlsx";
                buscaSheet.Title = tipo == "S"
                    ? "Selecione a Planilha Modelo de Sócios"
                    : "Selecione a Planilha Modelo de Acerto Semanal";

                if (buscaSheet.ShowDialog() == DialogResult.OK)
                {
                    return sheet.SaveExcelToDatabase(buscaSheet.FileName, tipo, empresaId) == 0;
                }
            }

            return false;
        }

        public bool AlteraNomeDoSocio(string filePath, string nome, string mes, string tipo, int ano)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Planilha não encontrada.\n\nFavor tentar novamente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        MessageBox.Show("A planilha não possui nenhuma aba.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    // EPPlus usa índice 0-based para Worksheets (0 = primeira aba, ex.: Janeiro / SEMANA 01)
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    if (tipo == "S")
                    {
                        // Altera o NOME (B2) e opcionalmente o ANO (J2) na Planilha de SÓCIOS
                        worksheet.Cells[2, 2].Value = nome;

                        if (ano > 0)
                            worksheet.Cells[2, 10].Value = ano;
                    }
                    else
                    {
                        // Altera o NOME (B3) e opcionalmente o MÊS (C7) na Planilha de ACERTO
                        worksheet.Cells[3, 2].Value = nome;

                        if (!string.IsNullOrWhiteSpace(mes))
                            worksheet.Cells[7, 3].Value = mes;
                    }

                    package.Save();
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar a planilha: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private string ConverteMes(int num)
        {
            switch (num)
            {
                case 1: return "JANEIRO";
                case 2: return "FEVEREIRO";
                case 3: return "MARÇO";
                case 4: return "ABRIL";
                case 5: return "MAIO";
                case 6: return "JUNHO";
                case 7: return "JULHO";
                case 8: return "AGOSTO";
                case 9: return "SETEMBRO";
                case 10: return "OUTUBRO";
                case 11: return "NOVEMBRO";
                case 12: return "DEZEMBRO";
                default: return "";
            }
        }
    }
}
