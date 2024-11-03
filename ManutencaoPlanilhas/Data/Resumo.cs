using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

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

        public void CarregaTabelaTemporaria(string filePath)
        {
            
        }
    }
}
