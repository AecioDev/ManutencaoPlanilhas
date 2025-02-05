﻿using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace ManutencaoPlanilhas.Data
{
    public class PlanilhasBD
    {
        //private string _connectionString = @"Data Source=Data\planilhas.db;Version=3;";
        private string _connectionString = @"Data Source=" + Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "planilhas.db") + ";Version=3;";


        public Int16 SaveExcelToDatabase(string filePath, string tipo)
        {
            try
            {
                // Verifica se o arquivo existe
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("O arquivo especificado não foi encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1; // Falha, arquivo não encontrado
                }

                // Lê o conteúdo do arquivo Excel
                byte[] fileBytes = File.ReadAllBytes(filePath);

                // Abre a conexão com o banco de dados SQLite
                using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
                {
                    conn.Open();
                    string query = "INSERT INTO Planilhas (NomeArquivo, ArquivoBlob, DataInclusao, Tipo) VALUES (@Nome, @Arquivo, @Data, @Tipo)";
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Nome", Path.GetFileName(filePath));
                        cmd.Parameters.AddWithValue("@Arquivo", fileBytes);
                        cmd.Parameters.AddWithValue("@Data", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        cmd.Parameters.AddWithValue("@Tipo", tipo);

                        cmd.ExecuteNonQuery(); // Executa o comando de inserção
                    }
                }

                //MessageBox.Show("Planilha salva com sucesso no banco de dados.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0; // Sucesso
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar a planilha no banco de dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2; // Falha
            }
        }

        public int RetrieveExcelFromDatabase(string tipo, string destinationPath)
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
                {
                    conn.Open();
                    string query = "SELECT ArquivoBlob FROM Planilhas WHERE Tipo = @Tipo ORDER BY DataInclusao DESC LIMIT 1";
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Tipo", tipo);

                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader["ArquivoBlob"] != DBNull.Value)
                                {
                                    byte[] excelData = (byte[])reader["ArquivoBlob"];
                                    File.WriteAllBytes(destinationPath, excelData);
                                    //MessageBox.Show("Planilha recuperada com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return 0; // Sucesso
                                }
                                else
                                {
                                    MessageBox.Show("BLOB vazio ou nulo encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return 1; // Falha, BLOB vazio
                                }
                            }
                            else
                            {
                                MessageBox.Show("Não foi encontrada a planilha de Modelo.\n\nFavor Selecione a Planilha de Modelo na próxima Tela para contiuar o processo...", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return 1; // Falha, nenhum resultado
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao recuperar a planilha: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2; // Falha
            }
        }

        public void SaveConfig(string key, string value)
        {
            try
            {
                // Define o caminho para a pasta Data na raiz do sistema
                string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");

                // Cria a pasta Data se ela não existir
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // Caminho completo do arquivo de configurações
                string filePath = Path.Combine(folderPath, "config.txt");

                // Verifica se a chave já existe e atualiza ou adiciona a chave e o valor
                string[] lines = File.Exists(filePath) ? File.ReadAllLines(filePath) : new string[0];
                bool keyExists = false;

                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    foreach (string line in lines)
                    {
                        if (line.StartsWith($"{key}:"))
                        {
                            writer.WriteLine($"{key}: {value}"); // Atualiza a chave existente
                            keyExists = true;
                        }
                        else
                        {
                            writer.WriteLine(line); // Mantém as outras configurações
                        }
                    }

                    // Se a chave não existir, adiciona-a ao arquivo
                    if (!keyExists)
                    {
                        writer.WriteLine($"{key}: {value}");
                    }
                }

                MessageBox.Show("Configuração salva com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar a configuração: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetConfig(string key)
        {
            try
            {
                // Define o caminho para a pasta Data na raiz do sistema
                string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
                string filePath = Path.Combine(folderPath, "config.txt");

                // Verifica se o arquivo de configuração existe
                if (!File.Exists(filePath))
                {
                    //MessageBox.Show("Arquivo de configuração não encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }

                // Lê o arquivo linha por linha e busca a chave solicitada
                string[] lines = File.ReadAllLines(filePath);
                foreach (string line in lines)
                {
                    if (line.StartsWith($"{key}:"))
                    {
                        // Retorna o valor após a chave (removendo a parte 'key: ')
                        return line.Substring(key.Length + 2).Trim();
                    }
                }

                //MessageBox.Show($"Configuração '{key}' não encontrada.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao ler a configuração: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        
    }
}
