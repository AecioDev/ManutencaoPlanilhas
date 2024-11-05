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

                // Verificar e criar a tabela Acerto
                if (!TableExists(connection, "Acerto"))
                {
                    CreateAcertoTable(connection);
                }

                // Verificar e criar a tabela DespesasAcerto
                if (!TableExists(connection, "DespesasAcerto"))
                {
                    CreateDespesasAcertoTable(connection);
                }

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

    private bool TableExists(SQLiteConnection connection, string tableName)
    {
        using (var command = new SQLiteCommand($"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName}';", connection))
        {
            return command.ExecuteScalar() != null;
        }
    }

    private void CreateAcertoTable(SQLiteConnection connection)
    {
        string createTableQuery = @"
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

        using (var command = new SQLiteCommand(createTableQuery, connection))
        {
            command.ExecuteNonQuery();
        }
    }

    private void CreateDespesasAcertoTable(SQLiteConnection connection)
    {
        string createTableQuery = @"
            CREATE TABLE DespesasAcerto (
                despesasId INTEGER PRIMARY KEY AUTOINCREMENT,
                acertoId INTEGER,
                despesaTipo TEXT,
                despesaDescricao TEXT,
                despesaValor REAL,
                FOREIGN KEY (acertoId) REFERENCES Acerto(acertoId) ON DELETE CASCADE
            );";

        using (var command = new SQLiteCommand(createTableQuery, connection))
        {
            command.ExecuteNonQuery();
        }
    }
}
