using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeLicence
{
    class Services
    {

        public static string filePath =  LerCaminhoDoArquivo(); //LerCaminhoDoArquivo();

        // Função para ler o caminho do arquivo do config.txt
        public static string LerCaminhoDoArquivo()
        {
            try
            {
                // Obtém o diretório onde o executável está rodando
                string pastaAplicacao = AppDomain.CurrentDomain.BaseDirectory;
                string configPath = Path.Combine(pastaAplicacao, "config.txt");

                if (File.Exists(configPath))
                {
                    string caminhoArquivo = File.ReadAllText(configPath).Trim();
                    if (!string.IsNullOrWhiteSpace(caminhoArquivo) && File.Exists(caminhoArquivo))
                    {
                        return caminhoArquivo;
                    }
                }

                //MessageBox.Show("Arquivo de configuração não encontrado ou inválido. Selecione um arquivo Excel.");
                return ""; // Retorna uma string vazia caso não encontre o caminho válido
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao ler o arquivo de configuração: " + ex.Message);
                return "";
            }
        }



        public static void AdicionarNaPlanilha(
                string usuario,
                string dispositivo,
                string setor,
                string email,
                string password,
                string validade)
            {
                try
                {
                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var planilha = workbook.Worksheet("licenca");

                        // Encontra a próxima linha vazia
                        int ultimaLinha = planilha.LastRowUsed()?.RowNumber() ?? 1;
                        int novaLinha = ultimaLinha + 1;
                        //Usuario/Dispositivo/Setor/Licenca/password/validade
                        // Adiciona os dados na planilha
                        planilha.Cell(novaLinha, 1).Value = usuario;    // Coluna A
                        planilha.Cell(novaLinha, 2).Value = dispositivo; // Coluna B
                        planilha.Cell(novaLinha, 3).Value = setor;    // Coluna C
                        planilha.Cell(novaLinha, 4).Value = email;
                        planilha.Cell(novaLinha, 5).Value = password;
                        planilha.Cell(novaLinha, 6).Value = validade;

                        // Salva a planilha
                        workbook.Save();
                    }

                    //MessageBox.Show("Dados adicionados com sucesso!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao adicionar dados: " + ex.Message);
                }
            }

            public static DataTable ObterDadosExcel()
            {
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("Usuário");
                dataTable.Columns.Add("Dispositivo");
                dataTable.Columns.Add("Setor");
                dataTable.Columns.Add("Licença");
                dataTable.Columns.Add("Password");
                dataTable.Columns.Add("Validade");

                try
                {
                    if (!File.Exists(filePath))
                        return dataTable; // Retorna vazio se o arquivo não existir

                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var planilha = workbook.Worksheet("licenca");

                        foreach (var linha in planilha.RowsUsed().Skip(1)) // Pula cabeçalho se existir
                        {
                            var usuario     = linha.Cell(1).Value.ToString();
                            var dispositivo = linha.Cell(2).Value.ToString();
                            var setor       = linha.Cell(3).Value.ToString();
                            var licenca     = linha.Cell(4).Value.ToString();
                            var password    = linha.Cell(5).Value.ToString();
                            var validadeCell = linha.Cell(6).Value;
                            var validade = DateTime.TryParse(validadeCell.ToString(), out DateTime data)
                                ? data.ToString("dd/MM/yyyy")
                                : "Sem Data"; // Define um valor padrão caso a célula esteja vazia

                            dataTable.Rows.Add(usuario, dispositivo,setor, licenca, password, validade);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erro ao carregar Excel: " + ex.Message);
                }

                return dataTable;
            }
    }
}



