using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;
using static OfficeLicence.Services;

namespace OfficeLicence
{
    public partial class ExtractorData : MaterialForm
    {

        private Dictionary<string, string> senhaPorEmail = new Dictionary<string, string>();

        public ExtractorData()
        {

            InitializeComponent();
       
            var materialSkinManager = MaterialSkinManager.Instance;
                materialSkinManager.AddFormToManage(this);
                materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
                materialSkinManager.ColorScheme = new ColorScheme(
                    Primary.Blue400, Primary.Blue500,
                    Primary.Blue500, Accent.LightBlue200,
                    TextShade.WHITE
                );

            if (File.Exists("config.txt"))
            {
                textCaminho.Text = Services.filePath;
            }

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.DataSource = Services.ObterDadosExcel();
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit; // Associa o evento para salvar após edição

            this.dataGridView1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseClick);
            this.copiarToolStripMenuItem.Click += new System.EventHandler(this.copiarToolStripMenuItem_Click);

            AtualizarDataGridView();
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textUsuario.Text) &&
                !string.IsNullOrWhiteSpace(textDispositivo.Text) &&
                !string.IsNullOrWhiteSpace(textSetor.Text) &&
                !string.IsNullOrWhiteSpace(textPassword.Text) &&
                !string.IsNullOrWhiteSpace(textEmail.Text) &&
                !string.IsNullOrWhiteSpace(textValidade.Text))
            {
                Services.AdicionarNaPlanilha(textUsuario.Text, textDispositivo.Text, textSetor.Text, textEmail.Text, textPassword.Text, textValidade.Text);

                //limpar campos
                textUsuario.Text = null;
                textDispositivo.Text = null;
                textSetor.Text = null;
                textPassword.Text = null;
                textEmail.Text = null;
                textValidade.Text = null;

                MessageBox.Show("Dados Atualizado");

                AtualizarDataGridView();
            }
            else
            {
                MessageBox.Show("Por favor, preencha todos os campos.");
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.ReadOnly = false;
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["Password"].Index) // Apenas para a coluna de senha
            {
                string licenca = dataGridView1.Rows[e.RowIndex].Cells["Licença"].Value?.ToString();
                string novaSenha = dataGridView1.Rows[e.RowIndex].Cells["Password"].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(licenca) && !string.IsNullOrWhiteSpace(novaSenha))
                {
                    // Atualiza todas as linhas no DataGridView
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells["Licença"].Value?.ToString() == licenca)
                        {
                            row.Cells["Password"].Value = novaSenha; // Define a nova senha para todas as linhas da mesma licença
                        }
                    }

                    // Agora salva a alteração no Excel
                    SalvarAlteracoesNoExcel(licenca, novaSenha);
                }
            }
        }

        private void SalvarAlteracoesNoExcel(string licenca, string novaSenha)
        {
            try
            {
                string filePath = Services.LerCaminhoDoArquivo(); // Obtém o caminho do arquivo
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("Erro: Caminho da planilha não encontrado.");
                    return;
                }

                using (var workbook = new XLWorkbook(filePath))
                {
                    var planilha = workbook.Worksheet("licenca");

                    // Percorre todas as linhas da planilha e altera a senha
                    foreach (var row in planilha.RangeUsed().RowsUsed())
                    {
                        if (row.Cell(4).GetValue<string>() == licenca) // Coluna 4 = "Licença"
                        {
                            row.Cell(5).Value = novaSenha; // Coluna 5 = "Password"
                        }
                    }

                    // Salva as alterações no Excel
                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar alterações no Excel: " + ex.Message);
            }
        }

        private void AtualizarDataGridView()
        {
            var dadosOriginais = Services.ObterDadosExcel();

            if (dadosOriginais == null || dadosOriginais.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum dado encontrado na planilha.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dataGridView1.DataSource = null;
                return;
            }

            // Atualiza todas as senhas para a mesma da última ocorrência
            var senhaPorEmail = new Dictionary<string, string>();

            foreach (DataRow row in dadosOriginais.Rows)
            {
                string licenca = row["Licença"].ToString();
                string password = row["Password"].ToString();

                if (!string.IsNullOrWhiteSpace(licenca))
                {
                    senhaPorEmail[licenca] = password; // Mantém a última senha encontrada
                }
            }

            // Atualiza todas as senhas no DataTable
            foreach (DataRow row in dadosOriginais.Rows)
            {
                string licenca = row["Licença"].ToString();
                if (!string.IsNullOrWhiteSpace(licenca) && senhaPorEmail.ContainsKey(licenca))
                {
                    row["Password"] = senhaPorEmail[licenca]; // Aplica a senha mais recente a todas as ocorrências
                }
            }

            // Exibe os dados atualizados no DataGridView
            dataGridView1.DataSource = dadosOriginais;
        }

        private void copiarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count > 0)
            {
                Clipboard.SetText(dataGridView1.SelectedCells[0].Value.ToString());
            }
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var hit = dataGridView1.HitTest(e.X, e.Y);
                if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
                {
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Selected = true;
                    contextMenuStrip1.Show(dataGridView1, e.Location);
                }
            }
        }

        private void materialLabel2_Click(object sender, EventArgs e)
        {

        }

        private void materialButton2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Arquivos Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string caminhoArquivo = openFileDialog.FileName;

                    // Obtém a pasta do arquivo Excel e cria o caminho para config.txt
                    string pastaArquivo = Path.GetDirectoryName(Application.ExecutablePath);
                    string caminhoConfig = Path.Combine(pastaArquivo, "config.txt");

                    // Salva o caminho do Excel no config.txt
                    File.WriteAllText(caminhoConfig, caminhoArquivo);

                    // Atualiza a variável filePath dinamicamente
                    Services.filePath = caminhoArquivo;

                    // Exibe o caminho no campo de texto
                    textCaminho.Text = caminhoArquivo;

                    AtualizarDataGridView();
                }
            }
        }

        #region redes sociais

        private void textCaminho_Click(object sender, EventArgs e)
        {
            textCaminho.ReadOnly = true;
        }

        private void linkedin_Click(object sender, EventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://www.linkedin.com/in/edinaelson-g-037633124/",
                UseShellExecute = true
            });
        }

        private void instagram_Click(object sender, EventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://www.instagram.com/edi_cifer/",
                UseShellExecute = true
            });
        }

        private void facebook_Click(object sender, EventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://www.facebook.com/edinaelson.santos.5/",
                UseShellExecute = true
            });
        }

        private void github_Click(object sender, EventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://github.com/Edinaelson",
                UseShellExecute = true
            });
        }
    }

    #endregion redes socias
}

