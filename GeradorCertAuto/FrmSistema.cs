using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GeradorCertAuto
{
    public partial class FrmSistema : Form
    {
        List<ModeloCertificado> CertificadosSelecionados = new List<ModeloCertificado>();

        public FrmSistema()
        {
            InitializeComponent();
            txtLocal.Text = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Certificados {DateTime.Now:yyyyMMdd}");
            CarregarListaCertificados();
            AtualizaQtdeModelosSelecionados();
        }

        private void CarregarListaCertificados()
        {
            chkbModelos.Items.Clear();
            var certificados = ConexaoBanco.GetModelos();
            foreach (var certificado in certificados)
            {
                chkbModelos.Items.Add(certificado, CertificadosSelecionados.Where(p => p.Id == certificado.Id).Any());
            }
        }

        private void ChkbModelos_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var modelo = chkbModelos.Items[e.Index] as ModeloCertificado;
            if (modelo == null) return;

            if (e.NewValue == CheckState.Checked)
            {
                if (!CertificadosSelecionados.Contains(modelo))
                    CertificadosSelecionados.Add(modelo);
            }
            else if (e.NewValue == CheckState.Unchecked)
            {
                CertificadosSelecionados.Remove(modelo);
            }
            AtualizaQtdeModelosSelecionados();
        }

        private void BtnSelecionarPasta_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Selecione a pasta desejada";
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtLocal.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnAbrirPasta_Click(object sender, EventArgs e)
        {
            string caminho = txtLocal.Text;
            if (!string.IsNullOrWhiteSpace(caminho))
            {
                bool abrirPasta = true;
                if (!System.IO.Directory.Exists(caminho))
                {
                    if (MessageBox.Show($"A pasta '{caminho}' não existe. Deseja criar a pasta?", "Pasta não encontrada", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        try
                        {
                            System.IO.Directory.CreateDirectory(caminho);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Erro ao criar a pasta: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            abrirPasta = false;
                        }
                    }
                    else
                    {
                        abrirPasta = false;
                    }
                }
                if (abrirPasta)
                {
                    System.Diagnostics.Process.Start("explorer.exe", caminho);
                }
            }
            else
            {
                MessageBox.Show("Pasta não encontrada ou caminho inválido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnGerarTab_Click(object sender, EventArgs e)
        {
            if (CertificadosSelecionados == null || CertificadosSelecionados.Count == 0)
            {
                MessageBox.Show("Selecione ao menos um modelo para gerar as abas.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            var workbook = excelApp.Workbooks.Add();

            // Remove a planilha padrão criada automaticamente
            while (workbook.Sheets.Count > 0)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1]).Delete();
            }

            AtualizaQtdeModelosSelecionados();
            AtualizaQtdeLinhas(0);
            
            foreach (var modelo in CertificadosSelecionados)
            {
                var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.Add();
                // Nome da aba (máx 31 caracteres, sem caracteres inválidos)
                string nomeAba = modelo.Nome.Replace(" ", "_");
                if (nomeAba.Length > 31) nomeAba = nomeAba.Substring(0, 31);
                worksheet.Name = nomeAba;

                // Obtém as colunas do modelo
                List<string> colunas;
                if (string.IsNullOrWhiteSpace(modelo.ColunasExcel))
                {
                    if (string.IsNullOrWhiteSpace(Properties.Settings.Default.ColunasExcel))
                    {
                        colunas = new List<string> { "Nome", "Cargo", "Empresa" };
                    }
                    else
                    {
                        colunas = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(Properties.Settings.Default.ColunasExcel);
                    }
                }
                else
                {
                    colunas = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(modelo.ColunasExcel);
                }

                // Cabeçalhos da tabela
                for (int i = 0; i < colunas.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = colunas[i];
                }
                AtualizaProgressoModelos(CertificadosSelecionados.IndexOf(modelo) + 1);
            }

            // Caminho para salvar o arquivo
            string pasta = txtLocal.Text;
            if (string.IsNullOrWhiteSpace(pasta) || !System.IO.Directory.Exists(pasta))
            {
                pasta = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Certificados {DateTime.Now:yyyyMMdd}");
            }
            string caminhoExcel = System.IO.Path.Combine(pasta, "TabelasModelos.xlsx");
            workbook.SaveAs(caminhoExcel);

            MessageBox.Show("Arquivo salvo em: " + caminhoExcel, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            AtualizaQtdeModelosSelecionados();
        }

        private void BtnGerarCert_Click(object sender, EventArgs e)
        {

        }

        private void BtnFechar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja realmente fechar?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Close();
            }
        }

        private void BtnAtualizarLista_Click(object sender, EventArgs e)
        {
            CarregarListaCertificados();
        }

        private void AtualizaQtdeModelosSelecionados()
        {
            pgbModelos.Minimum = 0;
            pgbModelos.Maximum = CertificadosSelecionados.Count > 0 ? CertificadosSelecionados.Count : 1;
            AtualizaProgressoModelos(0);
            AtualizaQtdeLinhas(0);
        }

        private void AtualizaProgressoModelos(int valor)
        {
            lblProgModelos.Text = $"{valor} de {CertificadosSelecionados.Count} modelos";
            pgbModelos.Value = valor;
        }

        private void AtualizaQtdeLinhas(int total)
        {
            pgbLinhas.Minimum = 0;
            pgbLinhas.Maximum = total > 0 ? total : 1;
            AtualizaProgressoLinhas(0, total);
        }

        private void AtualizaProgressoLinhas(int valor, int total)
        {
            lblProgLinhas.Text = total > 0 ? $"{valor} de {total} linhas" : "";
            pgbLinhas.Value = valor;
        }
    }
}
