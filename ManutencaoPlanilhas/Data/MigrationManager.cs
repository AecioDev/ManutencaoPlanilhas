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
            EnsureDatabaseDirectoryExists();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                CreateTablesDB(connection);     // Tabelas base do BD
                UpdateTablesDB(connection);     // Garante colunas adicionadas em versões posteriores
                CreateIndexes(connection);      // Índices

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

    private void EnsureDatabaseDirectoryExists()
    {
        var builder = new SQLiteConnectionStringBuilder(_connectionString);
        if (!string.IsNullOrWhiteSpace(builder.DataSource))
        {
            string dir = Path.GetDirectoryName(builder.DataSource);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
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
                NomeEmpresa TEXT,
                cellValorTotal TEXT,
                cellRecebimento TEXT,
                cellMercPagas TEXT,
                cellValSocio TEXT,
                cellValFirma TEXT
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

    private void UpdateTablesDB(SQLiteConnection connection)
    {
        if (!TableExists(connection, "Empresa")) return;

        string[] colunasEmpresa =
        {
            "cellValorTotal",
            "cellRecebimento",
            "cellMercPagas",
            "cellValSocio",
            "cellValFirma"
        };

        foreach (var coluna in colunasEmpresa)
        {
            if (!ColumnExiste(connection, "Empresa", coluna))
            {
                ExecuteQuery($"ALTER TABLE Empresa ADD COLUMN {coluna} TEXT;", connection);
            }
        }
    }

    private void CreateIndexes(SQLiteConnection connection)
    {
        if (!TableExists(connection, "Planilhas")) return;

        // Substitui índice legado sem EmpresaId por índice composto com EmpresaId
        if (IndexExists(connection, "IDX_Planilhas_TipoAndData"))
        {
            ExecuteQuery(@"DROP INDEX IF EXISTS ""IDX_Planilhas_TipoAndData"";", connection);
        }

        if (!IndexExists(connection, "IDX_Planilhas_EmpresaTipoData"))
        {
            string indexQuery = @"CREATE INDEX ""IDX_Planilhas_EmpresaTipoData"" ON ""Planilhas"" (""EmpresaId"", ""Tipo"", ""DataInclusao"" DESC);";
            ExecuteQuery(indexQuery, connection);
        }
    }

    private bool TableExists(SQLiteConnection connection, string tableName)
    {
        using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name=@name;", connection))
        {
            command.Parameters.AddWithValue("@name", tableName);
            return command.ExecuteScalar() != null;
        }
    }

    private bool IndexExists(SQLiteConnection connection, string indexName)
    {
        using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='index' AND name=@name;", connection))
        {
            command.Parameters.AddWithValue("@name", indexName);
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
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private void ExecuteQuery(string sqlQuery, SQLiteConnection connection)
    {
        using (var command = new SQLiteCommand(sqlQuery, connection))
        {
            command.ExecuteNonQuery();
        }
    }
}
