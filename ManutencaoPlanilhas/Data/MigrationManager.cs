using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

public class MigrationManager
{
    private readonly string _connectionString;

    public MigrationManager(string connectionString)
    {
        _connectionString = connectionString;
    }

    public bool RunMigrations()
    {
        try
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                CreateTablesDB(connection);     //Adicionar aqui as tabelas do BD
                CreateIndexes(connection);      //Adicionar aqui os indices
                //UpdateTablesDB(connection);   //Criar aqui novos campos

                connection.Close();
            }

            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erro ao executar migrações: {ex.Message}", "Erro de Migração", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }
    

    private void CreateTablesDB(SQLiteConnection connection)
    {
        string tableQuery = "";

        if (!TableExists(connection, "Empresa"))
        {
            tableQuery = @"
            CREATE TABLE Empresa (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                NomeEmpresa TEXT
            );";

            ExecuteQuery(tableQuery, connection);
        }

        if (!TableExists(connection, "Planilhas"))
        {
            tableQuery = @"
            CREATE TABLE Planilhas (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
	            EmpresaId INTEGER,
                NomeArquivo TEXT,
                ArquivoBlob BLOB,
                DataInclusao TEXT,
                Tipo TEXT,
                FOREIGN KEY (EmpresaId) REFERENCES Empresa(Id) ON DELETE CASCADE
            );";

            ExecuteQuery(tableQuery, connection);
        }

        if (!TableExists(connection, "Acerto"))
        {
            tableQuery = @"
            CREATE TABLE Acerto (
                acertoId INTEGER PRIMARY KEY AUTOINCREMENT,
                ano INTEGER,
                mes INTEGER,
                socio TEXT,
                QntMercEntregues INTEGER,
                ValTotEntregues REAL,
                ValTotDevolvidas REAL,
                ValTotVendidas REAL,
                QntTotFichas INTEGER,
                ValTotFichas REAL,
                ValTotReceitas REAL,
                ValTotDespesas REAL,
                ValTotSaldo REAL,
                IndiceParteFirma REAL,
                IndiceParteSocio REAL,
                ValParteFirma REAL,
                ValParteSocio REAL,
                ValMercEntregues REAL,
                ValMercNaoEntregues REAL,
                ValMercRetSalao REAL,
                ValMercDevolvidas REAL,
                ObsAcerto TEXT
            );";

            ExecuteQuery(tableQuery, connection);
        }

        if (!TableExists(connection, "DespesasAcerto"))
        {
            tableQuery = @"
            CREATE TABLE DespesasAcerto (
                despesasId INTEGER PRIMARY KEY AUTOINCREMENT,
                acertoId INTEGER,
                despesaTipo TEXT,
                despesaDescricao TEXT,
                despesaValor REAL,
                FOREIGN KEY (acertoId) REFERENCES Acerto(acertoId) ON DELETE CASCADE
            );";

            ExecuteQuery(tableQuery, connection);
        }
    }

    private void CreateIndexes(SQLiteConnection connection)
    {
        string indexQuery = "";

        if (TableExists(connection, "Planilhas") && !IndexExists(connection, "IDX_Planilhas_TipoAndData"))
        {
            indexQuery = @"CREATE UNIQUE INDEX ""IDX_Planilhas_TipoAndData"" ON ""Planilhas"" (""Tipo"", ""DataInclusao"" DESC);";
            ExecuteQuery(indexQuery, connection);
        }
    }

    /* De3scomentar qd precisar
    private void UpdateTablesDB(SQLiteConnection connection)
    {
        string columnQuery = "";

        //Vai colocando as Colunas Necessárias para ser criadas no BD
        if (!ColumnExiste(connection, "nomeTabela", "nomeColuna"))
        {
            columnQuery = "ALTER TABLE nomeTabela ADD COLUMN nomeColuna TEXT;";

            ExecuteQuery(columnQuery, connection);
        }
    }
    */

    private bool TableExists(SQLiteConnection connection, string tableName)
    {
        using (var command = new SQLiteCommand($"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName}';", connection))
        {
            return command.ExecuteScalar() != null;
        }
    }

    private bool IndexExists(SQLiteConnection connection, string indexName)
    {
        using (var command = new SQLiteCommand($"SELECT name FROM sqlite_master WHERE type='index' AND name='{indexName}';", connection))
        {
            return command.ExecuteScalar() != null;
        }
    }

    private bool ColumnExiste(SQLiteConnection connection, string tableName, string columnName)
    {
        using (var command = new SQLiteCommand($"PRAGMA table_info({tableName});", connection))
        {
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    string coluna = reader["name"].ToString();
                    if (coluna.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true; // A coluna existe
                    }
                }
            }
        }
        return false; // A coluna não existe
    }

    private void ExecuteQuery(string sqlQuery, SQLiteConnection connection)
    {
        using (var command = new SQLiteCommand(sqlQuery, connection))
        {
            command.ExecuteNonQuery();
        }
    }
}
