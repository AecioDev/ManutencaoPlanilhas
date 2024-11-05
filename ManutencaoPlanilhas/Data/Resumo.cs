using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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

        /// <summary>
        /// 
        /// Recebe o Caminho de uma planilha de Sócio e faz a extração dos dados para o banco de dados.
        /// 
        /// </summary>
        /// <param name="caminhoArquivoExcel"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Cria uma Planilha de Resumo no Caminho Recebido
        /// </summary>
        /// <param name="caminhoPlanilhas"></param>
        /// <returns></returns>
        public bool GeraPlanilhaResumo(string caminhoPlanilhas, string NomeEmpresa, int AnoMov)
        {
            string caminhoArquivo = caminhoPlanilhas + @"\#_RESUMO_"+ NomeEmpresa+"_"+AnoMov.ToString()+".xlsx";

            if(File.Exists( caminhoArquivo))
            {
                File.Delete( caminhoArquivo );
            }

            string[] arquivos = Directory.GetFiles(caminhoPlanilhas, "*.xlsx"); // Filtra apenas arquivos com extensão .xlsx

            if(arquivos.Length <= 0)
            {
                MessageBox.Show("Não Existem Planilhas de Sócios para gerar um Resumo!!!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                        
            foreach (string arquivoSocio in arquivos)
            {
                string nomeSocio = Path.GetFileNameWithoutExtension(arquivoSocio);
                
                for (int i = 1; i < 13; i++)
                {

                }
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage pacote = new ExcelPackage())
            {
                // Loop para criar uma planilha para cada mês
                for (int mes = 1; mes <= 12; mes++)
                {
                    string nomeMes = new DateTime(AnoMov, mes, 1).ToString("MMMM"); // Nome do mês
                    ExcelWorksheet planilha = pacote.Workbook.Worksheets.Add(nomeMes);

                    planilha.Cells["A1:G1"].Merge = true;

                    planilha.Cells["A1"].Value = NomeEmpresa;       // Nome da Empresa
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
                    planilha.Cells["B4"].Value = "Valor Total";
                    planilha.Cells["C4"].Value = "Recebimento";
                    planilha.Cells["D4"].Value = "Percentual";
                    planilha.Cells["E4"].Value = "Mercadorias Pagas";
                    planilha.Cells["F4"].Value = "Sócio Ganhou";
                    planilha.Cells["G4"].Value = "Firma Ganhou";

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

                    /*
                    // Preenche as despesas (linhas) na primeira coluna
                    for (int i = 0; i < despesas.Length; i++)
                    {
                        planilha.Cells[i + 2, 1].Value = despesas[i];
                    }

                    // Exemplo de valores: insira os valores desejados
                    Random rand = new Random();
                    for (int linha = 2; linha <= despesas.Length + 1; linha++)
                    {
                        for (int coluna = 2; coluna <= cachorros.Length + 1; coluna++)
                        {
                            planilha.Cells[linha, coluna].Value = rand.Next(50, 200); // Valores aleatórios de exemplo
                        }
                    }

                    // Totalizador para cada cachorro
                    int linhaTotal = despesas.Length + 2;
                    planilha.Cells[linhaTotal, 1].Value = "Total";
                    for (int coluna = 2; coluna <= cachorros.Length + 1; coluna++)
                    {
                        planilha.Cells[linhaTotal, coluna].Formula = $"SUM({planilha.Cells[2, coluna].Address}:{planilha.Cells[despesas.Length + 1, coluna].Address})";
                    }
                    */

                    // Ajusta a largura das colunas
                    planilha.Cells[planilha.Dimension.Address].AutoFitColumns();
                }

                // Salva o arquivo Excel no caminho especificado
                FileInfo arquivoInfo = new FileInfo(caminhoArquivo);
                pacote.SaveAs(arquivoInfo);

                MessageBox.Show("Planilha de Resumo criada com sucesso.");
            }

            return true;
        }

        public void CriaPlanilhaCachorros(string caminhoPlanilhas)
        {
            // Define o caminho e o nome do arquivo Excel a ser criado
            string caminhoArquivo = caminhoPlanilhas + @"\GastosCachorros.xlsx";

            // Configura o contexto de licença do EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Nomes dos cachorros e das despesas
            string[] cachorros = { "Suzi", "Pitoco", "Cacau", "Jade" };
            string[] despesas = { "Ração", "Banho", "Tosa", "Remédios" };

            using (ExcelPackage pacote = new ExcelPackage())
            {
                // Loop para criar uma planilha para cada mês
                for (int mes = 1; mes <= 12; mes++)
                {
                    string nomeMes = new DateTime(2023, mes, 1).ToString("MMMM"); // Nome do mês
                    ExcelWorksheet planilha = pacote.Workbook.Worksheets.Add(nomeMes);

                    // Títulos das colunas
                    planilha.Cells[1, 1].Value = "Despesas";
                    for (int i = 0; i < cachorros.Length; i++)
                    {
                        planilha.Cells[1, i + 2].Value = cachorros[i];
                    }

                    // Preenche as despesas (linhas) na primeira coluna
                    for (int i = 0; i < despesas.Length; i++)
                    {
                        planilha.Cells[i + 2, 1].Value = despesas[i];
                    }

                    // Exemplo de valores: insira os valores desejados
                    Random rand = new Random();
                    for (int linha = 2; linha <= despesas.Length + 1; linha++)
                    {
                        for (int coluna = 2; coluna <= cachorros.Length + 1; coluna++)
                        {
                            planilha.Cells[linha, coluna].Value = rand.Next(50, 200); // Valores aleatórios de exemplo
                        }
                    }

                    // Totalizador para cada cachorro
                    int linhaTotal = despesas.Length + 2;
                    planilha.Cells[linhaTotal, 1].Value = "Total";
                    for (int coluna = 2; coluna <= cachorros.Length + 1; coluna++)
                    {
                        planilha.Cells[linhaTotal, coluna].Formula = $"SUM({planilha.Cells[2, coluna].Address}:{planilha.Cells[despesas.Length + 1, coluna].Address})";
                    }

                    // Ajusta a largura das colunas
                    planilha.Cells[planilha.Dimension.Address].AutoFitColumns();
                }

                // Salva o arquivo Excel no caminho especificado
                FileInfo arquivoInfo = new FileInfo(caminhoArquivo);
                pacote.SaveAs(arquivoInfo);
                                
                MessageBox.Show("Planilha dos Cachorros criada com sucesso.");
            }
        }

        /// <summary>
        /// Recebe um nome de mes e retorna o seu número correspondente 
        /// </summary>
        /// <param name="nomeMes"></param>
        /// <returns></returns>
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

    }
}
