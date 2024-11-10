using System;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static OfficeOpenXml.ExcelErrorValue;

namespace ManutencaoPlanilhas.Data
{
    public class Resumo
    {
        //Processo para gerar a Planilha de Resumo das Planilhas Filhas
        //
        // 1 - Ler os dados das planilhas Filhas
        // 2 - Gravar na tabela temporária
        // 3 - Ler a Tabela temporária
        // 4 - Gerar a Planilha de Resumo
        // 5 - Adicionar os valores de cada nome na planilha

        private readonly string _connectionString;

        public Resumo(string connectionString)
        {
            _connectionString = connectionString;
        }
        
        public string[] ObterPosicaoValoresResumoSocios(int empresaId)
        {
            string[] posicoes = new string[6];

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                string query = @"
                    SELECT NomeEmpresa, cellValorTotal, cellRecebimento, cellMercPagas, cellValSocio, cellValFirma
                    FROM Empresa
                    WHERE Id = @EmpresaId";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@EmpresaId", empresaId);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            posicoes[0] = reader.IsDBNull(0) ? "" : reader.GetString(0); // NomeEmpresa
                            posicoes[1] = reader.IsDBNull(1) ? "" : reader.GetString(1); // cellValorTotal
                            posicoes[2] = reader.IsDBNull(2) ? "" : reader.GetString(2); // cellRecebimento
                            posicoes[3] = reader.IsDBNull(3) ? "" : reader.GetString(3); // cellMercPagas
                            posicoes[4] = reader.IsDBNull(4) ? "" : reader.GetString(4); // cellValSocio
                            posicoes[5] = reader.IsDBNull(5) ? "" : reader.GetString(5); // cellValFirma
                        }
                    }
                }
            }

            return posicoes;
        }
                
        public bool GeraPlanilhaResumoSocios(string caminhoPlanilhas, int empresaId, int AnoMov)
        {
            const int LinhaInicial = 5;
            int linha;
            int LinhaFinal;
            int LinhaTotal;
            string[] cellValores = ObterPosicaoValoresResumoSocios(empresaId);
            string caminhoArquivo = caminhoPlanilhas + @"\#_RESUMO_" + cellValores[0] + "_" + AnoMov.ToString() + ".xlsx";

            if (File.Exists(caminhoArquivo))
            {
                File.Delete(caminhoArquivo);
            }

            string[] arquivos = Directory.GetFiles(caminhoPlanilhas, "*.xlsx")
                             .Where(arquivo => !Path.GetFileName(arquivo).StartsWith("~"))
                             .ToArray();

            if (arquivos.Length > 0)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage pacote = new ExcelPackage())
                {
                    // Loop para criar uma planilha para cada mês
                    for (int mes = 1; mes <= 12; mes++)
                    {
                        string nomeMes = new DateTime(AnoMov, mes, 1).ToString("MMMM"); // Nome do mês
                        var nomePlanilha = CapitalizarPrimeiraLetra(nomeMes);
                        ExcelWorksheet planilha = pacote.Workbook.Worksheets.Add(nomePlanilha);

                        planilha.Cells["A1:G1"].Merge = true;

                        planilha.Cells["A1"].Value = cellValores[0];       // Nome da Empresa
                        planilha.Cells["A1"].Style.Font.Size = 26;      // Tamanho da fonte 26
                        planilha.Cells["A1"].Style.Font.Bold = true;    // Negrito                    
                        planilha.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        planilha.Cells["A2:G2"].Merge = true;

                        planilha.Cells["A2"].Value = "RESULTADO DE GANHOS POR SÓCIO";
                        planilha.Cells["A2"].Style.Font.Size = 18;      // Tamanho da fonte 26
                        planilha.Cells["A2"].Style.Font.Bold = true;    // Negrito                    
                        planilha.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        //Mes [linha, coluna]
                        planilha.Cells["A3"].Value = nomeMes.ToUpper();
                        planilha.Cells["F3"].Value = "ANO:";
                        planilha.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        planilha.Cells["G3"].Value = AnoMov.ToString();

                        planilha.Cells["A3:G3"].Style.Font.Size = 18;      // Tamanho da fonte 26
                        planilha.Cells["A3:G3"].Style.Font.Bold = true;    // Negrito   

                        planilha.Row(4).Height = 34;

                        // Títulos das colunas
                        planilha.Cells["A4"].Value = "Sócio";
                        planilha.Column(1).Width = 30;

                        planilha.Cells["B4"].Value = "Valor Total";
                        planilha.Column(2).Width = 25;

                        planilha.Cells["C4"].Value = "Recebimento";
                        planilha.Column(3).Width = 25;

                        planilha.Cells["D4"].Value = "Percentual";
                        planilha.Column(4).Width = 20;

                        planilha.Cells["E4"].Value = "Mercadorias Pagas";
                        planilha.Column(5).Width = 25;

                        planilha.Cells["F4"].Value = "Sócio Ganhou";
                        planilha.Column(6).Width = 25;

                        planilha.Cells["G4"].Value = "Firma Ganhou";
                        planilha.Column(7).Width = 25;

                        planilha.Cells["A4:G4"].Style.Font.Size = 16;      // Tamanho da fonte 26
                        planilha.Cells["A4:G4"].Style.Font.Bold = true;    // Negrito   
                        planilha.Cells["A4:G4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        planilha.Cells["A4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        // Define bordas espessas para a linha 4, de A4 até G4
                        using (var range = planilha.Cells["A4:G4"])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Medium;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        }

                        linha = LinhaInicial;
                        
                        // Cria dos Campos de Formulas do Resumo.
                        foreach (string arquivoSocio in arquivos)
                        {
                            var nomeSocio = Path.GetFileNameWithoutExtension(arquivoSocio);
                            var nomeArquivo = Path.GetFileName(arquivoSocio);

                            //Cell [linha, coluna]
                            planilha.Row(linha).Height = 34;

                            //Nome do Sócio     A
                            planilha.Cells[linha, 1].Value = nomeSocio.ToUpper();
                            planilha.Cells[linha, 1].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Valor Total   B
                            planilha.Cells[linha, 2].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!" + cellValores[1].Trim();
                            planilha.Cells[linha, 2].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Recebimento   C
                            planilha.Cells[linha, 3].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!" + cellValores[2].Trim();
                            planilha.Cells[linha, 3].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Percentual    D
                            planilha.Cells[linha, 4].Formula = "=IF(B" + linha.ToString() + " > 0, C" + linha.ToString() + " / B" + linha.ToString() + ", 0)";
                            planilha.Cells[linha, 4].Style.Numberformat.Format = "0.00%";
                            planilha.Cells[linha, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Mercadorias Pagas E
                            planilha.Cells[linha, 5].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!" + cellValores[3].Trim();
                            planilha.Cells[linha, 5].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Sócio Ganhou  F
                            planilha.Cells[linha, 6].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!" + cellValores[4].Trim();
                            planilha.Cells[linha, 6].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Firma Ganhou  G
                            planilha.Cells[linha, 7].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!" + cellValores[5].Trim();
                            planilha.Cells[linha, 7].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 7].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 7].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            planilha.Cells[linha, 1, linha, 7].Style.Font.Size = 16;
                            planilha.Cells[linha, 1, linha, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            linha++;
                        }

                        //Borda Final
                        LinhaFinal = linha - 1;
                        planilha.Cells[LinhaFinal, 1, LinhaFinal, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                        LinhaTotal = LinhaFinal + 1;
                        for (int coluna = 2; coluna <= 7; coluna++)
                        {
                            if (coluna != 4)
                            {
                                planilha.Cells[LinhaTotal, coluna].Formula = $"SUM({planilha.Cells[LinhaInicial, coluna].Address}:{planilha.Cells[LinhaFinal, coluna].Address})";
                                planilha.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            }
                            else
                            {
                                planilha.Cells[LinhaTotal, coluna].Formula = "=IF(B" + LinhaTotal.ToString() + " > 0, C" + LinhaTotal.ToString() + " / B" + LinhaTotal.ToString() + ", 0)";
                                planilha.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "0.00%";
                            }
                        }

                        planilha.Cells[LinhaTotal, 2, LinhaTotal, 7].Style.Font.Size = 16;
                        planilha.Cells[LinhaTotal, 2, LinhaTotal, 7].Style.Font.Bold = true;
                        planilha.Cells[LinhaTotal, 2, LinhaTotal, 7].Style.Font.Color.SetColor(Color.Blue);
                        planilha.Cells[LinhaTotal, 2, LinhaTotal, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    }

                    // Salva o arquivo Excel no caminho especificado
                    FileInfo arquivoInfo = new FileInfo(caminhoArquivo);
                    pacote.SaveAs(arquivoInfo);

                    MessageBox.Show("Planilha de Resumo criada com sucesso.");
                }

                return true;
            }

            MessageBox.Show("Não Existem Planilhas de Sócios para gerar um Resumo!!!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }


        //Percorrer um laço para os meses do ano
        //Entrar na pasta do mes com nome: 01 - JANEIRO 
        //Buscar os Sócios dentro da pasta de Mes: 
        //Iniciar um laço para cada sócio e gravar as fórmulas na planilha.

        public bool GeraPlanilhaResumoAcerto(string caminhoPastas, int empresaId, int AnoMov)
        {
            const int LinhaInicial = 7;
            string[] cellValores = ObterPosicaoValoresResumoSocios(empresaId);
            int linha;
            int LinhaFinal;
            int LinhaTotal;

            // Loop para percorrer os Meses do Ano
            for (int mes = 1; mes <= 12; mes++)
            {
                var nomeMes = new DateTime(AnoMov, mes, 1).ToString("MMMM"); // Nome do mês
                var mesChar = mes < 10 ? "0" + mes.ToString() : mes.ToString();
                var nomePasta = mesChar + " - " + nomeMes.ToUpper(); //01 - JANEIRO 
                var caminhoPlanilhas = caminhoPastas + @"\" + nomePasta;

                var caminhoArquivo = caminhoPlanilhas + @"\#_RESUMO_ACERTO_SEMANAL_" + nomeMes.ToUpper() + "_" + AnoMov.ToString() + ".xlsx";

                if (File.Exists(caminhoArquivo))
                {
                    File.Delete(caminhoArquivo);
                }

                string[] arquivos = Directory.GetFiles(caminhoPlanilhas, "*.xlsx")
                                 .Where(arquivo => !Path.GetFileName(arquivo).StartsWith("~"))
                                 .ToArray();

                if (arquivos.Length <= 0)
                {
                    MessageBox.Show("Não Existem Planilhas de Sócios para gerar um Resumo!!!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage pacote = new ExcelPackage())
                {
                    // Loop para percorrer as SEMANAS do MES
                    for (int semana = 1; semana <= 4; semana++)
                    {
                        var semanaChar = "0" + semana.ToString();
                        var nomePlanilha = "SEMANA " + semanaChar; //SEMANA 01

                        ExcelWorksheet planilha = pacote.Workbook.Worksheets.Add(nomePlanilha);

                        planilha.Cells["A1:E1"].Merge = true;

                        planilha.Cells["A1"].Value = cellValores[0];    // Nome da Empresa
                        planilha.Cells["A1"].Style.Font.Size = 22;      // Tamanho da fonte 26
                        planilha.Cells["A1"].Style.Font.Bold = true;    // Negrito                    
                        
                        planilha.Cells["A2:E2"].Merge = true;

                        planilha.Cells["A2"].Value = "RESUMO DE ACERTO SEMANAL - " + nomePlanilha;
                        planilha.Cells["A2"].Style.Font.Size = 16;      // Tamanho da fonte
                        planilha.Cells["A2"].Style.Font.Bold = true;    // Negrito              

                        //Mes [linha, coluna]
                        planilha.Cells["A4"].Value = nomeMes.ToUpper();
                        planilha.Cells["B4"].Value = AnoMov.ToString();

                        planilha.Cells["A4:B4"].Style.Font.Size = 18;      // Tamanho da fonte 26
                        planilha.Cells["A4:B4"].Style.Font.Bold = true;    // Negrito   

                        //Tópicos do Resumo
                        planilha.Row(5).Height = 30;
                        planilha.Cells["A5:U5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        

                        planilha.Cells["A5:E5"].Merge = true;
                        planilha.Cells["A5"].Value = "Resumo";
                        planilha.Cells["A5:E5"].Style.Fill.BackgroundColor.SetColor(Color.Gray);    //Cor de Fundo
                        planilha.Cells["A5:E5"].Style.Border.BorderAround(ExcelBorderStyle.Medium); //Borda ao Redor

                        planilha.Cells["F5:N5"].Merge = true;
                        planilha.Cells["F5"].Value = "Despesas";
                        planilha.Cells["F5:N5"].Style.Fill.BackgroundColor.SetColor(Color.DarkRed);    //Cor de Fundo
                        planilha.Cells["F5:N5"].Style.Border.BorderAround(ExcelBorderStyle.Medium); //Borda ao Redor

                        planilha.Cells["O5:U5"].Merge = true;
                        planilha.Cells["O5"].Value = "Recebimentos";
                        planilha.Cells["O5:U5"].Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);  //Cor de Fundo
                        planilha.Cells["O5:U5"].Style.Border.BorderAround(ExcelBorderStyle.Medium); //Borda ao Redor

                        using (var range = planilha.Cells["A5:U5"])
                        {
                            // Formatação
                            range.Style.Font.Size = 16;                       // Tamanho da fonte 26
                            range.Style.Font.Bold = true;                     // Negrito   
                            range.Style.Font.Name = "Arial Rounded MT Bold";  //Tipo da Fonte
                            range.Style.Font.Color.SetColor(Color.White);     //Cor da Fonte
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        }
                        
                        // RESUMO
                        planilha.Cells["A6"].Value = "NOME";
                        planilha.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(1).Width = 20.7;

                        planilha.Cells["B6"].Value = "COBRANÇA";
                        planilha.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(2).Width = 14.7;

                        planilha.Cells["C6"].Value = "QTD NOTAS";
                        planilha.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(3).Width = 11.7;

                        planilha.Cells["D6"].Value = "RECEBIMENTO";
                        planilha.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(4).Width = 14.7;

                        planilha.Cells["E6"].Value = "%";
                        planilha.Cells["E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        planilha.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(5).Width = 7.7;

                        // DESPESAS
                        planilha.Cells["F6"].Value = "COMBUSTÍVEL";
                        planilha.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(6).Width = 13.7; 

                        planilha.Cells["G6"].Value = "REFEIÇÃO";
                        planilha.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(7).Width = 12.7;

                        planilha.Cells["H6"].Value = "HOTEL";
                        planilha.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(8).Width = 12.7;

                        planilha.Cells["I6"].Value = "LAVA JATO";
                        planilha.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(9).Width = 12.7;

                        planilha.Cells["J6"].Value = "OFICINA";
                        planilha.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(10).Width = 12.7;

                        planilha.Cells["K6"].Value = "FICHAS";
                        planilha.Cells["K6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(11).Width = 12.7;

                        planilha.Cells["L6"].Value = "VALE";
                        planilha.Cells["L6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(12).Width = 12.7;

                        planilha.Cells["M6"].Value = "VALE AJUDANTE";
                        planilha.Cells["M6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(13).Width = 15.7;

                        planilha.Cells["N6"].Value = "TOTAL DESPESAS";
                        planilha.Cells["N6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(14).Width = 15.7;

                        //RECEBIMENTOS
                        planilha.Cells["O6"].Value = "TOTAL A PAGAR";
                        planilha.Cells["O6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(15).Width = 15.7;

                        planilha.Cells["P6"].Value = "PIX";
                        planilha.Cells["P6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(16).Width = 12.7;

                        planilha.Cells["Q6"].Value = "DEPÓSITO";
                        planilha.Cells["Q6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(17).Width = 12.7;

                        planilha.Cells["R6"].Value = "DINHEIRO";
                        planilha.Cells["R6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(18).Width = 12.7;

                        planilha.Cells["S6"].Value = "CHEQUE";
                        planilha.Cells["S6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(19).Width = 12.7;

                        planilha.Cells["T6"].Value = "TOTAL PAGAMENTO";
                        planilha.Cells["T6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(20).Width = 19.7;

                        planilha.Cells["U6"].Value = "FALTA/SOBRA";
                        planilha.Cells["U6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        planilha.Column(21).Width = 14.7;

                        // Formatação
                        planilha.Cells["A6:U6"].Style.Font.Size = 11;      // Tamanho da fonte 26
                        planilha.Cells["A6:U6"].Style.Font.Bold = true;    // Negrito   

                        // Bloco de dados dos Sócios
                        linha = LinhaInicial;

                        // Cria dos Campos de Formulas do Resumo.
                        foreach (string arquivoSocio in arquivos)
                        {
                            var nomeSocio = Path.GetFileNameWithoutExtension(arquivoSocio);
                            var nomeArquivo = Path.GetFileName(arquivoSocio);

                            //Cell [linha, coluna]
                            //planilha.Row(linha).Height = 34;

                            /***** RESUMO *****/

                            //SOCIO     A
                            planilha.Cells[linha, 1].Value = nomeSocio.ToUpper();
                            planilha.Cells[linha, 1].Style.Font.Bold = true;
                            planilha.Cells[linha, 1].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //COBRANÇA  B
                            planilha.Cells[linha, 2].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$4";
                            planilha.Cells[linha, 2].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //QTD NOTAS C
                            planilha.Cells[linha, 3].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$E$4";
                            planilha.Cells[linha, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            planilha.Cells[linha, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //RECEBIMENTO   D
                            planilha.Cells[linha, 4].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$6";
                            planilha.Cells[linha, 4].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //PERCENTUAL    E
                            planilha.Cells[linha, 5].Formula = $"IF(B{linha} > 0, D{linha}/B{linha}, 0)";
                            planilha.Cells[linha, 5].Style.Numberformat.Format = "0.00%";
                            planilha.Cells[linha, 5].Style.Font.Bold = true;
                            planilha.Cells[linha, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 5].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            /***** DESPESAS *****/

                            //COMBUSTÍVEL   F
                            planilha.Cells[linha, 6].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$9";
                            planilha.Cells[linha, 6].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 6].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //REFEIÇÃO      G
                            planilha.Cells[linha, 7].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$10";
                            planilha.Cells[linha, 7].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 7].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 7].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //HOTEL         H
                            planilha.Cells[linha, 8].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$11";
                            planilha.Cells[linha, 8].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 8].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //LAVA JATO     I
                            planilha.Cells[linha, 9].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$12";
                            planilha.Cells[linha, 9].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //OFICINA       J
                            planilha.Cells[linha, 10].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$13";
                            planilha.Cells[linha, 10].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //FICHAS        K
                            planilha.Cells[linha, 11].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$14";
                            planilha.Cells[linha, 11].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //VALE          L
                            planilha.Cells[linha, 12].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$16";
                            planilha.Cells[linha, 12].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 12].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 12].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //VALE AJUDANTE M
                            planilha.Cells[linha, 13].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$15";
                            planilha.Cells[linha, 13].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 13].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 13].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //TOTAL DESPESAS    N
                            planilha.Cells[linha, 14].Formula = $"SUM(F{linha}:M{linha})";
                            planilha.Cells[linha, 14].Style.Font.Bold = true;
                            planilha.Cells[linha, 14].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 14].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                            /***** RECEBIMENTOS *****/

                            //TOTAL A PAGAR     O
                            planilha.Cells[linha, 15].Formula = $"D{linha}-N{linha}";
                            planilha.Cells[linha, 15].Style.Font.Bold = true;
                            planilha.Cells[linha, 15].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 15].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //PIX               P
                            planilha.Cells[linha, 16].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$21";
                            planilha.Cells[linha, 16].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 16].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 16].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 16].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //DEPÓSITO          Q
                            planilha.Cells[linha, 17].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$22";
                            planilha.Cells[linha, 17].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 17].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 17].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 17].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //DINHEIRO          R
                            planilha.Cells[linha, 18].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$23";
                            planilha.Cells[linha, 18].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 18].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 18].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 18].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //CHEQUE            S
                            planilha.Cells[linha, 19].Formula = "='" + caminhoPlanilhas + @"\[" + nomeArquivo + "]" + nomePlanilha + "'!$C$24";
                            planilha.Cells[linha, 19].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 19].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 19].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 19].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //TOTAL PAGAMENTO   T
                            planilha.Cells[linha, 20].Formula = $"SUM(P{linha}:S{linha})";
                            planilha.Cells[linha, 20].Style.Font.Bold = true;
                            planilha.Cells[linha, 20].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 20].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 20].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 20].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //FALTA/SOBRA       U
                            planilha.Cells[linha, 21].Formula = $"T{linha}-O{linha}";
                            planilha.Cells[linha, 21].Style.Font.Bold = true;
                            planilha.Cells[linha, 21].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            planilha.Cells[linha, 21].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            planilha.Cells[linha, 21].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            planilha.Cells[linha, 21].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //Formatação
                            planilha.Cells[linha, 1, linha, 21].Style.Font.Size = 11;

                            linha++;
                        } // For Sócios

                        //Valores Totais

                        LinhaFinal = linha - 1;
                        LinhaTotal = LinhaFinal + 1;
                        for (int coluna = 1; coluna <= 21; coluna++)
                        {
                            if (coluna != 1 && coluna != 5 && coluna != 15 && coluna != 21)
                            {                                
                                planilha.Cells[LinhaTotal, coluna].Formula = $"SUM({planilha.Cells[LinhaInicial, coluna].Address}:{planilha.Cells[LinhaFinal, coluna].Address})";
                                planilha.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";                                
                            }
                            else
                            {
                                if (coluna == 5) //Percentual $"IF(B{linha} > 0, D{linha}/B{linha}, 0)";
                                {                                    
                                    planilha.Cells[LinhaTotal, coluna].Formula = $"IF(B{LinhaTotal} > 0, D{LinhaTotal}/B{LinhaTotal}, 0)";
                                    planilha.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "0.00%";
                                }
                                else if (coluna == 15)
                                {
                                    planilha.Cells[LinhaTotal, coluna].Formula = $"D{LinhaTotal}-N{LinhaTotal}";
                                    planilha.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                                }
                                else if (coluna == 21)
                                {
                                    planilha.Cells[LinhaTotal, coluna].Formula = $"T{LinhaTotal}-O{LinhaTotal}";
                                    planilha.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                                }
                                else
                                {
                                    planilha.Cells[LinhaTotal, coluna].Value = "TOTAIS";
                                }
                                
                            }

                            planilha.Cells[LinhaTotal, coluna].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        }

                        // Formatação Linha Total
                        using (var range = planilha.Cells[LinhaTotal, 1, LinhaTotal, 21])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;

                            range.Style.Font.Size = 11;                     // Tamanho da fonte 26
                            range.Style.Font.Bold = true;                   // Negrito
                            range.Style.Font.Color.SetColor(Color.Yellow);  //Cor da Fonte
                        }
                                                
                        planilha.Cells[LinhaTotal, 1, LinhaTotal, 5].Style.Fill.BackgroundColor.SetColor(Color.Gray);   //Formatação Totais Resumo
                        planilha.Cells[LinhaTotal, 6, LinhaTotal, 14].Style.Fill.BackgroundColor.SetColor(Color.DarkRed);   //Formatação Totais Despesas
                        planilha.Cells[LinhaTotal, 15, LinhaTotal, 21].Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);   //Formatação Totais Recebimentos

                    } // For da Semanas 1 a 4 

                    // Planilha de TOTAIS DO MÊS

                    var nomePlanilhaFinal = "TOTAL DO MÊS";
                    ExcelWorksheet planilhaFinal = pacote.Workbook.Worksheets.Add(nomePlanilhaFinal);

                    planilhaFinal.Cells["A1:E1"].Merge = true;

                    planilhaFinal.Cells["A1"].Value = cellValores[0];    // Nome da Empresa
                    planilhaFinal.Cells["A1"].Style.Font.Size = 22;      // Tamanho da fonte 26
                    planilhaFinal.Cells["A1"].Style.Font.Bold = true;    // Negrito                    

                    planilhaFinal.Cells["A2:E2"].Merge = true;

                    planilhaFinal.Cells["A2"].Value = "RESUMO DE ACERTO SEMANAL - " + nomePlanilhaFinal;
                    planilhaFinal.Cells["A2"].Style.Font.Size = 16;      // Tamanho da fonte
                    planilhaFinal.Cells["A2"].Style.Font.Bold = true;    // Negrito              

                    //Mes [linha, coluna]
                    planilhaFinal.Cells["A4"].Value = nomeMes.ToUpper();
                    planilhaFinal.Cells["B4"].Value = AnoMov.ToString();

                    planilhaFinal.Cells["A4:B4"].Style.Font.Size = 18;      // Tamanho da fonte 26
                    planilhaFinal.Cells["A4:B4"].Style.Font.Bold = true;    // Negrito   

                    //Tópicos do Resumo
                    planilhaFinal.Row(5).Height = 30;
                    planilhaFinal.Cells["A5:U5"].Style.Fill.PatternType = ExcelFillStyle.Solid;


                    planilhaFinal.Cells["A5:E5"].Merge = true;
                    planilhaFinal.Cells["A5"].Value = "Resumo";
                    planilhaFinal.Cells["A5:E5"].Style.Fill.BackgroundColor.SetColor(Color.Gray);    //Cor de Fundo
                    planilhaFinal.Cells["A5:E5"].Style.Border.BorderAround(ExcelBorderStyle.Medium); //Borda ao Redor

                    planilhaFinal.Cells["F5:N5"].Merge = true;
                    planilhaFinal.Cells["F5"].Value = "Despesas";
                    planilhaFinal.Cells["F5:N5"].Style.Fill.BackgroundColor.SetColor(Color.DarkRed);    //Cor de Fundo
                    planilhaFinal.Cells["F5:N5"].Style.Border.BorderAround(ExcelBorderStyle.Medium); //Borda ao Redor

                    planilhaFinal.Cells["O5:U5"].Merge = true;
                    planilhaFinal.Cells["O5"].Value = "Recebimentos";
                    planilhaFinal.Cells["O5:U5"].Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);  //Cor de Fundo
                    planilhaFinal.Cells["O5:U5"].Style.Border.BorderAround(ExcelBorderStyle.Medium); //Borda ao Redor

                    using (var range = planilhaFinal.Cells["A5:U5"])
                    {
                        // Formatação
                        range.Style.Font.Size = 16;                       // Tamanho da fonte 26
                        range.Style.Font.Bold = true;                     // Negrito   
                        range.Style.Font.Name = "Arial Rounded MT Bold";  //Tipo da Fonte
                        range.Style.Font.Color.SetColor(Color.White);     //Cor da Fonte
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    }

                    // RESUMO
                    planilhaFinal.Cells["A6"].Value = "NOME";
                    planilhaFinal.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(1).Width = 20.7;

                    planilhaFinal.Cells["B6"].Value = "COBRANÇA";
                    planilhaFinal.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(2).Width = 14.7;

                    planilhaFinal.Cells["C6"].Value = "QTD NOTAS";
                    planilhaFinal.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(3).Width = 11.7;

                    planilhaFinal.Cells["D6"].Value = "RECEBIMENTO";
                    planilhaFinal.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(4).Width = 14.7;

                    planilhaFinal.Cells["E6"].Value = "%";
                    planilhaFinal.Cells["E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaFinal.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(5).Width = 7.7;

                    // DESPESAS
                    planilhaFinal.Cells["F6"].Value = "COMBUSTÍVEL";
                    planilhaFinal.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(6).Width = 13.7;

                    planilhaFinal.Cells["G6"].Value = "REFEIÇÃO";
                    planilhaFinal.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(7).Width = 12.7;

                    planilhaFinal.Cells["H6"].Value = "HOTEL";
                    planilhaFinal.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(8).Width = 12.7;

                    planilhaFinal.Cells["I6"].Value = "LAVA JATO";
                    planilhaFinal.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(9).Width = 12.7;

                    planilhaFinal.Cells["J6"].Value = "OFICINA";
                    planilhaFinal.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(10).Width = 12.7;

                    planilhaFinal.Cells["K6"].Value = "FICHAS";
                    planilhaFinal.Cells["K6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(11).Width = 12.7;

                    planilhaFinal.Cells["L6"].Value = "VALE";
                    planilhaFinal.Cells["L6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(12).Width = 12.7;

                    planilhaFinal.Cells["M6"].Value = "VALE AJUDANTE";
                    planilhaFinal.Cells["M6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(13).Width = 15.7;

                    planilhaFinal.Cells["N6"].Value = "TOTAL DESPESAS";
                    planilhaFinal.Cells["N6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(14).Width = 15.7;

                    //RECEBIMENTOS
                    planilhaFinal.Cells["O6"].Value = "TOTAL A PAGAR";
                    planilhaFinal.Cells["O6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(15).Width = 15.7;

                    planilhaFinal.Cells["P6"].Value = "PIX";
                    planilhaFinal.Cells["P6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(16).Width = 12.7;

                    planilhaFinal.Cells["Q6"].Value = "DEPÓSITO";
                    planilhaFinal.Cells["Q6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(17).Width = 12.7;

                    planilhaFinal.Cells["R6"].Value = "DINHEIRO";
                    planilhaFinal.Cells["R6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(18).Width = 12.7;

                    planilhaFinal.Cells["S6"].Value = "CHEQUE";
                    planilhaFinal.Cells["S6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(19).Width = 12.7;

                    planilhaFinal.Cells["T6"].Value = "TOTAL PAGAMENTO";
                    planilhaFinal.Cells["T6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(20).Width = 19.7;

                    planilhaFinal.Cells["U6"].Value = "FALTA/SOBRA";
                    planilhaFinal.Cells["U6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    planilhaFinal.Column(21).Width = 14.7;

                    // Formatação
                    planilhaFinal.Cells["A6:U6"].Style.Font.Size = 11;      // Tamanho da fonte 26
                    planilhaFinal.Cells["A6:U6"].Style.Font.Bold = true;    // Negrito   

                    // Bloco de dados dos Sócios
                    linha = LinhaInicial;

                    // Cria dos Campos de Formulas do Resumo.
                    foreach (string arquivoSocio in arquivos)
                    {
                        var nomeSocio = Path.GetFileNameWithoutExtension(arquivoSocio);
                        var nomeArquivo = Path.GetFileName(arquivoSocio);

                        //Cell [linha, coluna]
                        //planilhaFinal.Row(linha).Height = 34;

                        /***** RESUMO *****/

                        //SOCIO     A
                        planilhaFinal.Cells[linha, 1].Value = nomeSocio.ToUpper();
                        planilhaFinal.Cells[linha, 1].Style.Font.Bold = true;
                        planilhaFinal.Cells[linha, 1].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //COBRANÇA  B
                        planilhaFinal.Cells[linha, 2].Formula = $"'SEMANA 01'!B{linha}+'SEMANA 02'!B{linha}+'SEMANA 03'!B{linha}+'SEMANA 04'!B{linha}";
                        planilhaFinal.Cells[linha, 2].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //QTD NOTAS C
                        planilhaFinal.Cells[linha, 3].Formula = $"'SEMANA 01'!C{linha}+'SEMANA 02'!C{linha}+'SEMANA 03'!C{linha}+'SEMANA 04'!C{linha}";
                        planilhaFinal.Cells[linha, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        planilhaFinal.Cells[linha, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //RECEBIMENTO   D
                        planilhaFinal.Cells[linha, 4].Formula = $"'SEMANA 01'!D{linha}+'SEMANA 02'!D{linha}+'SEMANA 03'!D{linha}+'SEMANA 04'!D{linha}";
                        planilhaFinal.Cells[linha, 4].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //PERCENTUAL    E
                        planilhaFinal.Cells[linha, 5].Formula = $"IF(B{linha} > 0, D{linha}/B{linha}, 0)";
                        planilhaFinal.Cells[linha, 5].Style.Numberformat.Format = "0.00%";
                        planilhaFinal.Cells[linha, 5].Style.Font.Bold = true;
                        planilhaFinal.Cells[linha, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 5].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        /***** DESPESAS *****/

                        //COMBUSTÍVEL   F
                        planilhaFinal.Cells[linha, 6].Formula = $"'SEMANA 01'!F{linha}+'SEMANA 02'!F{linha}+'SEMANA 03'!F{linha}+'SEMANA 04'!F{linha}";
                        planilhaFinal.Cells[linha, 6].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 6].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //REFEIÇÃO      G
                        planilhaFinal.Cells[linha, 7].Formula = $"'SEMANA 01'!G{linha}+'SEMANA 02'!G{linha}+'SEMANA 03'!G{linha}+'SEMANA 04'!G{linha}";
                        planilhaFinal.Cells[linha, 7].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 7].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 7].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //HOTEL         H
                        planilhaFinal.Cells[linha, 8].Formula = $"'SEMANA 01'!H{linha}+'SEMANA 02'!H{linha}+'SEMANA 03'!H{linha}+'SEMANA 04'!H{linha}";
                        planilhaFinal.Cells[linha, 8].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 8].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //LAVA JATO     I
                        planilhaFinal.Cells[linha, 9].Formula = $"'SEMANA 01'!I{linha}+'SEMANA 02'!I{linha}+'SEMANA 03'!I{linha}+'SEMANA 04'!I{linha}";
                        planilhaFinal.Cells[linha, 9].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //OFICINA       J
                        planilhaFinal.Cells[linha, 10].Formula = $"'SEMANA 01'!J{linha}+'SEMANA 02'!J{linha}+'SEMANA 03'!J{linha}+'SEMANA 04'!J{linha}";
                        planilhaFinal.Cells[linha, 10].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //FICHAS        K
                        planilhaFinal.Cells[linha, 11].Formula = $"'SEMANA 01'!K{linha}+'SEMANA 02'!K{linha}+'SEMANA 03'!K{linha}+'SEMANA 04'!K{linha}";
                        planilhaFinal.Cells[linha, 11].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //VALE          L
                        planilhaFinal.Cells[linha, 12].Formula = $"'SEMANA 01'!L{linha}+'SEMANA 02'!L{linha}+'SEMANA 03'!L{linha}+'SEMANA 04'!L{linha}";
                        planilhaFinal.Cells[linha, 12].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 12].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 12].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //VALE AJUDANTE M
                        planilhaFinal.Cells[linha, 13].Formula = $"'SEMANA 01'!M{linha}+'SEMANA 02'!M{linha}+'SEMANA 03'!M{linha}+'SEMANA 04'!M{linha}";
                        planilhaFinal.Cells[linha, 13].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 13].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 13].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //TOTAL DESPESAS    N
                        planilhaFinal.Cells[linha, 14].Formula = $"SUM(F{linha}:M{linha})";
                        planilhaFinal.Cells[linha, 14].Style.Font.Bold = true;
                        planilhaFinal.Cells[linha, 14].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 14].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                        /***** RECEBIMENTOS *****/

                        //TOTAL A PAGAR     O
                        planilhaFinal.Cells[linha, 15].Formula = $"D{linha}-N{linha}";
                        planilhaFinal.Cells[linha, 15].Style.Font.Bold = true;
                        planilhaFinal.Cells[linha, 15].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 15].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //PIX               P
                        planilhaFinal.Cells[linha, 16].Formula = $"'SEMANA 01'!P{linha}+'SEMANA 02'!P{linha}+'SEMANA 03'!P{linha}+'SEMANA 04'!P{linha}";
                        planilhaFinal.Cells[linha, 16].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 16].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 16].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 16].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //DEPÓSITO          Q
                        planilhaFinal.Cells[linha, 17].Formula = $"'SEMANA 01'!Q{linha}+'SEMANA 02'!Q{linha}+'SEMANA 03'!Q{linha}+'SEMANA 04'!Q{linha}";
                        planilhaFinal.Cells[linha, 17].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 17].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 17].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 17].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //DINHEIRO          R
                        planilhaFinal.Cells[linha, 18].Formula = $"'SEMANA 01'!R{linha}+'SEMANA 02'!R{linha}+'SEMANA 03'!R{linha}+'SEMANA 04'!R{linha}";
                        planilhaFinal.Cells[linha, 18].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 18].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 18].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 18].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //CHEQUE            S
                        planilhaFinal.Cells[linha, 19].Formula = $"'SEMANA 01'!S{linha}+'SEMANA 02'!S{linha}+'SEMANA 03'!S{linha}+'SEMANA 04'!S{linha}";
                        planilhaFinal.Cells[linha, 19].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 19].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 19].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 19].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //TOTAL PAGAMENTO   T
                        planilhaFinal.Cells[linha, 20].Formula = $"SUM(P{linha}:S{linha})";
                        planilhaFinal.Cells[linha, 20].Style.Font.Bold = true;
                        planilhaFinal.Cells[linha, 20].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 20].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 20].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 20].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //FALTA/SOBRA       U
                        planilhaFinal.Cells[linha, 21].Formula = $"T{linha}-O{linha}";
                        planilhaFinal.Cells[linha, 21].Style.Font.Bold = true;
                        planilhaFinal.Cells[linha, 21].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        planilhaFinal.Cells[linha, 21].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        planilhaFinal.Cells[linha, 21].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        planilhaFinal.Cells[linha, 21].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        //Formatação
                        planilhaFinal.Cells[linha, 1, linha, 21].Style.Font.Size = 11;

                        linha++;
                    } // For Sócios

                    //Valores Totais

                    LinhaFinal = linha - 1;
                    LinhaTotal = LinhaFinal + 1;
                    for (int coluna = 1; coluna <= 21; coluna++)
                    {
                        if (coluna != 1 && coluna != 5 && coluna != 15 && coluna != 21)
                        {
                            planilhaFinal.Cells[LinhaTotal, coluna].Formula = $"SUM({planilhaFinal.Cells[LinhaInicial, coluna].Address}:{planilhaFinal.Cells[LinhaFinal, coluna].Address})";
                            planilhaFinal.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                        }
                        else
                        {
                            if (coluna == 5) //Percentual $"IF(B{linha} > 0, D{linha}/B{linha}, 0)";
                            {
                                planilhaFinal.Cells[LinhaTotal, coluna].Formula = $"IF(B{LinhaTotal} > 0, D{LinhaTotal}/B{LinhaTotal}, 0)";
                                planilhaFinal.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "0.00%";
                            }
                            else if (coluna == 15)
                            {
                                planilhaFinal.Cells[LinhaTotal, coluna].Formula = $"D{LinhaTotal}-N{LinhaTotal}";
                                planilhaFinal.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            }
                            else if (coluna == 21)
                            {
                                planilhaFinal.Cells[LinhaTotal, coluna].Formula = $"T{LinhaTotal}-O{LinhaTotal}";
                                planilhaFinal.Cells[LinhaTotal, coluna].Style.Numberformat.Format = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * \"-\"??_-;_-@";
                            }
                            else
                            {
                                planilhaFinal.Cells[LinhaTotal, coluna].Value = "TOTAIS";
                            }

                        }

                        planilhaFinal.Cells[LinhaTotal, coluna].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    }

                    // Formatação Linha Total
                    using (var range = planilhaFinal.Cells[LinhaTotal, 1, LinhaTotal, 21])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;

                        range.Style.Font.Size = 11;                     // Tamanho da fonte 26
                        range.Style.Font.Bold = true;                   // Negrito
                        range.Style.Font.Color.SetColor(Color.Yellow);  //Cor da Fonte
                    }

                    planilhaFinal.Cells[LinhaTotal, 1, LinhaTotal, 5].Style.Fill.BackgroundColor.SetColor(Color.Gray);   //Formatação Totais Resumo
                    planilhaFinal.Cells[LinhaTotal, 6, LinhaTotal, 14].Style.Fill.BackgroundColor.SetColor(Color.DarkRed);   //Formatação Totais Despesas
                    planilhaFinal.Cells[LinhaTotal, 15, LinhaTotal, 21].Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);   //Formatação Totais Recebimentos

                    // FIM planilhaFinal TOTAL DO MÊS

                    // Salva o arquivo Excel no caminho especificado
                    FileInfo arquivoInfo = new FileInfo(caminhoArquivo);
                    pacote.SaveAs(arquivoInfo);

                } //Using Excel Package

            } //For Meses do Ano

            return true;
        }
        
        private int ObterMesPorNome(string nomeMes)
        {
            int mesNum = -1;

            switch(nomeMes.ToLower())
            {
                case "janeiro":
                    mesNum = 1;
                    break;
                case "fevereiro":
                    mesNum = 2;
                    break;
                case "março":
                    mesNum = 3;
                    break;
                case "abril":
                    mesNum = 4;
                    break;
                case "maio":
                    mesNum = 5;
                    break;
                case "junho":
                    mesNum = 6;
                    break;
                case "julho":
                    mesNum = 7;
                    break;
                case "agosto":
                    mesNum = 8;
                    break;
                case "setembro":
                    mesNum = 9;
                    break;
                case "outubro":
                    mesNum = 10;
                    break;
                case "novembro":
                    mesNum = 11;
                    break;
                case "dezembro":
                    mesNum = 12;
                    break;
                default:
                    break;
            }

            return mesNum;
        }

        public string CapitalizarPrimeiraLetra(string texto)
        {
            if (string.IsNullOrEmpty(texto))
                return texto;

            return char.ToUpper(texto[0]) + texto.Substring(1).ToLower();
        }

        public bool CarregaTabelaTemporaria(string caminhoArquivoExcel)
        {
            try
            {
                if (!File.Exists(caminhoArquivoExcel))
                {
                    MessageBox.Show("Arquivo de planilha não encontrado.");
                    return false;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(caminhoArquivoExcel)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        long acertoId = 0;
                        string nomeMes = worksheet.Name; // Obtem o nome da planilha (ex: "Janeiro", "Fevereiro", etc.)
                        int mes = ObterMesPorNome(nomeMes);

                        if (mes == -1)
                        {
                            MessageBox.Show($"Nome de planilha não reconhecido: {nomeMes}");
                            continue;
                        }

                        // Extrai dados das células da planilha
                        int ano = Convert.ToInt32(worksheet.Cells["J2"].Value);
                        string socio = worksheet.Cells["B2"].Text;
                        int qntMercEntregues = Convert.ToInt32(worksheet.Cells["C6"].Value);
                        decimal valTotEntregues = Convert.ToDecimal(worksheet.Cells["D6"].Value);
                        decimal valTotDevolvidas = Convert.ToDecimal(worksheet.Cells["E6"].Value);
                        decimal valTotVendidas = Convert.ToDecimal(worksheet.Cells["F6"].Value); /*valTotEntregues - valTotDevolvidas*/
                        int qntTotFichas = Convert.ToInt32(worksheet.Cells["B11"].Value);
                        decimal valTotFichas = Convert.ToDecimal(worksheet.Cells["B12"].Value);
                        decimal valTotReceitas = Convert.ToDecimal(worksheet.Cells["D11"].Value);
                        decimal valTotDespesas = Convert.ToDecimal(worksheet.Cells["D12"].Value);
                        decimal valTotSaldo = Convert.ToDecimal(worksheet.Cells["D13"].Value);   /*valTotReceitas - valTotDespesas*/
                        decimal indiceParteFirma = Convert.ToDecimal(worksheet.Cells["D19"].Value);
                        decimal indiceParteSocio = Convert.ToDecimal(worksheet.Cells["D20"].Value);
                        decimal valParteFirma = Convert.ToDecimal(worksheet.Cells["E19"].Value);
                        decimal valParteSocio = Convert.ToDecimal(worksheet.Cells["E20"].Value);
                        decimal valMercEntregues = Convert.ToDecimal(worksheet.Cells["J7"].Value);
                        decimal valMercNaoEntregues = Convert.ToDecimal(worksheet.Cells["J9"].Value);
                        decimal valMercRetSalao = Convert.ToDecimal(worksheet.Cells["J11"].Value);
                        decimal ValMercDevolvidas = Convert.ToDecimal(worksheet.Cells["J13"].Value); /*valMercEntregues + valMercNaoEntregues - valMercRetSalao*/
                        string obsAcerto = worksheet.Cells["C24"].Text;

                        using (var connection = new SQLiteConnection(_connectionString))
                        {
                            connection.Open();
                            // Executa a inserção no banco de dados para cada planilha/mês

                            string insertQuery = @"
                            INSERT INTO Acerto (
                                ano, mes, socio, QntMercEntregues, ValTotEntregues, ValTotDevolvidas, ValTotVendidas, 
                                QntTotFichas, ValTotFichas, ValTotReceitas, ValTotDespesas, ValTotSaldo, 
                                IndiceParteFirma, IndiceParteSocio, ValParteFirma, ValParteSocio, 
                                ValMercEntregues, ValMercNaoEntregues, ValMercRetSalao, ValMercDevolvidas, ObsAcerto
                            ) VALUES (
                                @ano, @mes, @socio, @QntMercEntregues, @ValTotEntregues, @ValTotDevolvidas, @ValTotVendidas, 
                                @QntTotFichas, @ValTotFichas, @ValTotReceitas, @ValTotDespesas, @ValTotSaldo, 
                                @IndiceParteFirma, @IndiceParteSocio, @ValParteFirma, @ValParteSocio, 
                                @ValMercEntregues, @ValMercNaoEntregues, @ValMercRetSalao, @ValMercDevolvidas, @ObsAcerto
                            );
                            SELECT last_insert_rowid();"; // Pega o ID do último registro inserido para associar a outras tabelas";

                            using (var command = new SQLiteCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@ano", ano);
                                command.Parameters.AddWithValue("@mes", mes);
                                command.Parameters.AddWithValue("@socio", socio);
                                command.Parameters.AddWithValue("@QntMercEntregues", qntMercEntregues);
                                command.Parameters.AddWithValue("@ValTotEntregues", valTotEntregues);
                                command.Parameters.AddWithValue("@ValTotDevolvidas", valTotDevolvidas);
                                command.Parameters.AddWithValue("@ValTotVendidas", valTotVendidas);
                                command.Parameters.AddWithValue("@QntTotFichas", qntTotFichas);
                                command.Parameters.AddWithValue("@ValTotFichas", valTotFichas);
                                command.Parameters.AddWithValue("@ValTotReceitas", valTotReceitas);
                                command.Parameters.AddWithValue("@ValTotDespesas", valTotDespesas);
                                command.Parameters.AddWithValue("@ValTotSaldo", valTotSaldo);
                                command.Parameters.AddWithValue("@IndiceParteFirma", indiceParteFirma);
                                command.Parameters.AddWithValue("@IndiceParteSocio", indiceParteSocio);
                                command.Parameters.AddWithValue("@ValParteFirma", valParteFirma);
                                command.Parameters.AddWithValue("@ValParteSocio", valParteSocio);
                                command.Parameters.AddWithValue("@ValMercEntregues", valMercEntregues);
                                command.Parameters.AddWithValue("@ValMercNaoEntregues", valMercNaoEntregues);
                                command.Parameters.AddWithValue("@ValMercRetSalao", valMercRetSalao);
                                command.Parameters.AddWithValue("@ValMercDevolvidas", ValMercDevolvidas);
                                command.Parameters.AddWithValue("@ObsAcerto", obsAcerto);

                                acertoId = (long)command.ExecuteScalar(); // Recupera o ID do novo acerto
                            }


                            // Extrai dados das despesas da planilha
                            // Loop por cada despesa, dependendo da estrutura de onde as despesas estão na planilha
                            string despesaTipo = "";
                            string despesaDescricao = "";
                            decimal despesaValor = 0;
                            string cellDescricao = "";
                            string cellValor = "";

                            for (int i = 13; i < 30; i++)
                            {
                                if (i == 13)
                                {
                                    despesaTipo = "DESPESAS";
                                    continue;
                                }

                                if (i == 18)
                                {
                                    despesaTipo = "OFICINAS";
                                    continue;
                                }

                                if (i == 24)
                                {
                                    despesaTipo = "OUTROS";
                                    continue;
                                }

                                cellDescricao = "A" + i.ToString();
                                cellValor = "B" + i.ToString();

                                despesaDescricao = worksheet.Cells[cellDescricao].Text;
                                despesaValor = Convert.ToDecimal(worksheet.Cells[cellValor].Value);


                                if (despesaValor > 0)
                                {

                                    // Executa a inserção no banco de dados para cada despesa

                                    string queryDespesas = @"
                                    INSERT INTO DespesasAcerto 
                                    (acertoId, despesaTipo, despesaDescricao, despesaValor) 
                                    VALUES (@acertoId, @despesaTipo, @despesaDescricao, @despesaValor);";

                                    using (var commandDespesas = new SQLiteCommand(queryDespesas, connection))
                                    {
                                        commandDespesas.Parameters.AddWithValue("@acertoId", acertoId);
                                        commandDespesas.Parameters.AddWithValue("@despesaTipo", despesaTipo);
                                        commandDespesas.Parameters.AddWithValue("@despesaDescricao", despesaDescricao);
                                        commandDespesas.Parameters.AddWithValue("@despesaValor", despesaValor);
                                        commandDespesas.ExecuteNonQuery();
                                    }

                                }
                            } //For Despesas

                            connection.Close();

                        } //using sqlite

                    } //foreach meses

                    package.Dispose();

                    MessageBox.Show("Dados da planilha salvos com sucesso.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar dados da planilha: {ex.Message}");
                return false;
            }
        }
    }
}
