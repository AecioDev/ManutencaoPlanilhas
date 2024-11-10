using System.Collections.Generic;
using System.Data.SQLite;
using System;
using System.IO;

namespace ManutencaoPlanilhas.Data
{
    public class Empresa
    {
        public int Id { get; set; }
        public string NomeEmpresa { get; set; }

        public override string ToString()
        {
            return NomeEmpresa; // Exibe o nome da empresa na ComboBox
        }
    }

    public class EmpresaDB 
    { 
        //private string _connectionString = @"Data Source=Data\planilhas.db;Version=3;";
        private string _connectionString = @"Data Source=" + Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "planilhas.db") + ";Version=3;";

        public List<Empresa> ObterEmpresas()
        {
            List<Empresa> empresas = new List<Empresa>();

            using (SQLiteConnection connection = new SQLiteConnection(_connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "SELECT Id, NomeEmpresa FROM Empresa";

                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    {
                        using (SQLiteDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                empresas.Add(new Empresa
                                {
                                    Id = reader.GetInt32(0),
                                    NomeEmpresa = reader.GetString(1)
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Tratar o erro de acordo com a necessidade
                    Console.WriteLine("Erro ao obter empresas: " + ex.Message);
                }
            }

            return empresas;
        }
    }
}
