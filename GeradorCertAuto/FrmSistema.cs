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
            txtLocal.Text = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Certificados{DateTime.Now:yyyyMMdd}");
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
                if (!CertificadosSelecionados.Where(p => p.Id == modelo.Id).Any())
                    CertificadosSelecionados.Add(modelo);
            }
            else if (e.NewValue == CheckState.Unchecked)
            {
                if (CertificadosSelecionados.Where(p => p.Id == modelo.Id).Any())
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
            try
            {
                TravarBotoes(true);
                if (CertificadosSelecionados == null || CertificadosSelecionados.Count == 0)
                {
                    MessageBox.Show("Selecione ao menos um modelo para gerar as abas.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TravarBotoes(false);
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
                if (string.IsNullOrWhiteSpace(pasta))
                {
                    pasta = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Certificados{DateTime.Now:yyyyMMdd}");
                    txtLocal.Text = pasta;
                }
                else if (!System.IO.Directory.Exists(pasta))
                {
                    System.IO.Directory.CreateDirectory(pasta);
                }

                string caminhoExcel = System.IO.Path.Combine(pasta, "TabelasModelos.xlsx");
                workbook.SaveAs(caminhoExcel);

                MessageBox.Show("Arquivo salvo em: " + caminhoExcel, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AtualizaQtdeModelosSelecionados();
                TravarBotoes(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao gerar a tabela: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TravarBotoes(false);
            }
        }

        private void BtnGerarCert_Click(object sender, EventArgs e)
        {
            try
            {
                TravarBotoes(true);

                if (CertificadosSelecionados == null || CertificadosSelecionados.Count == 0)
                {
                    MessageBox.Show("Selecione ao menos um modelo para gerar os certificados.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TravarBotoes(false);
                    return;
                }

                string pasta = txtLocal.Text;
                if (string.IsNullOrWhiteSpace(pasta) || !System.IO.Directory.Exists(pasta))
                {
                    MessageBox.Show("Selecione uma pasta válida para salvar os certificados.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TravarBotoes(false);
                    return;
                }

                string caminhoExcel = System.IO.Path.Combine(pasta, "TabelasModelos.xlsx");
                if (!System.IO.File.Exists(caminhoExcel))
                {
                    MessageBox.Show("A planilha TabelasModelos.xlsx não foi encontrada. Gere a tabela antes de criar os certificados.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TravarBotoes(false);
                    return;
                }

                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                var workbook = excelApp.Workbooks.Open(caminhoExcel);

                AtualizaQtdeModelosSelecionados();
                int modeloIndex = 0;

                foreach (var modelo in CertificadosSelecionados)
                {
                    modeloIndex++;
                    // Tenta encontrar a aba pelo nome do modelo
                    string nomeAba = modelo.Nome.Replace(" ", "_");
                    if (nomeAba.Length > 31) nomeAba = nomeAba.Substring(0, 31);

                    Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Sheets)
                    {
                        if (ws.Name == nomeAba)
                        {
                            worksheet = ws;
                            break;
                        }
                    }
                    if (worksheet == null)
                        continue; // Não encontrou a aba, pula para o próximo modelo

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

                    // Descobre o número de linhas preenchidas
                    int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;

                    AtualizaProgressoModelos(modeloIndex);
                    AtualizaQtdeLinhas(lastRow - 1);

                    // Garante que a pasta do modelo existe
                    string pastaModelo = System.IO.Path.Combine(pasta, nomeAba);
                    if (!System.IO.Directory.Exists(pastaModelo))
                        System.IO.Directory.CreateDirectory(pastaModelo);

                    for (int i = 2; i <= lastRow; i++) // Começa na linha 2 (dados)
                    {
                        // Lê os valores das colunas para a linha atual
                        var valores = new Dictionary<string, string>();
                        for (int c = 0; c < colunas.Count; c++)
                        {
                            var valorCelula = worksheet.Cells[i, c + 1].Text;
                            valores[colunas[c]] = valorCelula?.ToString() ?? string.Empty;
                        }

                        // Gera o certificado usando o modelo
                        // 1. Salva o arquivo temporário do modelo
                        string tempPptx = System.IO.Path.Combine(pastaModelo, $"_temp_{Guid.NewGuid()}.pptx");
                        modelo.SalvarArquivo(tempPptx);

                        // 2. Abre a apresentação
                        var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                        var novaApresentacao = pptApp.Presentations.Open(tempPptx, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                        // 3. Substitui palavras-chave em todos os slides
                        foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in novaApresentacao.Slides)
                        {
                            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                            {
                                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                                    shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    var tr = shape.TextFrame.TextRange;
                                    foreach (var coluna in colunas)
                                    {
                                        if (tr.Text.Contains($"<<{coluna}>>") || tr.Text.Contains($"<<{coluna}Maiusculo>>") || tr.Text.Contains($"<<{coluna}PriMaiusculo>>"))
                                        {
                                            var valor = valores[coluna].Trim();
                                            SubstituirTextoPreservandoFormatacao(tr, $"<<{coluna}>>", valor);
                                            SubstituirTextoPreservandoFormatacao(tr, $"<<{coluna}Maiusculo>>", valor.ToUpper());
                                            SubstituirTextoPreservandoFormatacao(tr, $"<<{coluna}PriMaiusculo>>", PrimeirasMaisculas(valor));
                                        }
                                    }
                                }
                            }
                        }

                        // 4. Salva o certificado
                        string nomeArquivoBase = valores[colunas[0]];
                        string nomeArquivoPptx = System.IO.Path.Combine(pastaModelo, $"{PrimeiroUltimoNome(nomeArquivoBase).Replace(" ", "")}.pptx");
                        int contador = 0;
                        while (System.IO.File.Exists(nomeArquivoPptx))
                        {
                            contador++;
                            nomeArquivoPptx = System.IO.Path.Combine(pastaModelo, $"{PrimeiroUltimoNome(nomeArquivoBase).Replace(" ", "")} {contador}.pptx");
                        }
                        novaApresentacao.SaveAs(nomeArquivoPptx, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                        // 5. Salva PDF
                        string nomeArquivoPdf = nomeArquivoPptx.Replace(".pptx", ".pdf");
                        novaApresentacao.SaveAs(nomeArquivoPdf, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF);

                        novaApresentacao.Close();
                        System.IO.File.Delete(tempPptx);

                        AtualizaProgressoLinhas(i - 1, lastRow - 1);
                    }
                }

                workbook.Close(false);
                excelApp.Quit();

                MessageBox.Show("Certificados gerados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AtualizaQtdeModelosSelecionados();
                TravarBotoes(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao gerar os certificados: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TravarBotoes(false);
            }
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
            Application.DoEvents();
        }

        private void AtualizaProgressoModelos(int valor)
        {
            lblProgModelos.Text = $"{valor} de {CertificadosSelecionados.Count} modelos";
            pgbModelos.Value = valor;
            Application.DoEvents();
        }

        private void AtualizaQtdeLinhas(int total)
        {
            pgbLinhas.Minimum = 0;
            pgbLinhas.Maximum = total > 0 ? total : 1;
            AtualizaProgressoLinhas(0, total);
            Application.DoEvents();
        }

        private void AtualizaProgressoLinhas(int valor, int total)
        {
            lblProgLinhas.Text = total > 0 ? $"{valor} de {total} linhas" : "";
            pgbLinhas.Value = valor;
            Application.DoEvents();
        }

        private void TravarBotoes(bool travar)
        {
            btnGerarTab.Enabled = !travar;
            btnGerarCert.Enabled = !travar;
            btnAbrirPasta.Enabled = !travar;
            btnSelecionarPasta.Enabled = !travar;
            btnAtualizarLista.Enabled = !travar;
            chkbModelos.Enabled = !travar;
            txtLocal.Enabled = !travar;
            Application.DoEvents();
        }

        // Funções auxiliares (copie do seu Gerador.cs ou adapte conforme necessário)
        private string PrimeiroUltimoNome(string nomeCompleto)
        {
            if (string.IsNullOrEmpty(nomeCompleto))
                return string.Empty;
            var partes = nomeCompleto.Trim().Split(' ');
            if (partes.Length == 0)
                return string.Empty;
            return PrimeirasMaisculas($"{partes[0]} {partes[partes.Length - 1]}");
        }

        private string PrimeirasMaisculas(string texto)
        {
            if (string.IsNullOrEmpty(texto))
                return string.Empty;
            return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(texto.Trim().ToLower());
        }

        private void SubstituirTextoPreservandoFormatacao(Microsoft.Office.Interop.PowerPoint.TextRange tr, string chave, string valor)
        {
            var encontrado = tr.Find(chave, 0, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            while (encontrado != null)
            {
                encontrado.Text = valor;
                encontrado = tr.Find(chave, 0, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            }
        }
    }
}
