using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GeradorCertAuto
{
    public partial class Gerador
    {
        private void Gerador_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnGerarExcel_Click(object sender, RibbonControlEventArgs e)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            var workbook = excelApp.Workbooks.Add();
            var worksheet = (Excel.Worksheet)workbook.Sheets[1];

            // Obtém os nomes das colunas da configuração
            var json = Properties.Settings.Default.ColunasExcel;
            List<string> colunas;
            if (string.IsNullOrWhiteSpace(json))
                colunas = new List<string> { "Nome", "Cargo", "Empresa" };
            else
                colunas = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(json);

            // Cabeçalhos da tabela
            for (int i = 0; i < colunas.Count; i++)
            {
                worksheet.Cells[1, i + 1] = colunas[i];
            }

            // Obter o caminho da apresentação ativa do PowerPoint
            var pptApp = Globals.ThisAddIn.Application;
            var presentation = pptApp.ActivePresentation;
            string pasta = System.IO.Path.GetDirectoryName(presentation.FullName);

            // Montar caminho do arquivo Excel
            string caminhoExcel = System.IO.Path.Combine(pasta, "TabelaAlunos.xlsx");
            workbook.SaveAs(caminhoExcel);

            System.Windows.Forms.MessageBox.Show("Arquivo salvo em: " + caminhoExcel);
        }

        private void BtnCriarCert_Click(object sender, RibbonControlEventArgs e)
        {
            var pptApp = Globals.ThisAddIn.Application;
            var presentation = pptApp.ActivePresentation;
            string pastaRaiz = System.IO.Path.GetDirectoryName(presentation.FullName);
            string pasta = System.IO.Path.Combine(pastaRaiz, "Certificados");
            if (!System.IO.Directory.Exists(pasta))
            {
                System.IO.Directory.CreateDirectory(pasta);
            }
            string caminhoExcel = System.IO.Path.Combine(pastaRaiz, "TabelaAlunos.xlsx");

            GerarCertificados(caminhoExcel, pasta);
        }
        private void GerarCertificados(string caminhoExcel, string pastaDestino)
        {
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Open(caminhoExcel);
            var worksheet = (Excel.Worksheet)workbook.Sheets[1];

            // Obtém os nomes das colunas da configuração
            var json = Properties.Settings.Default.ColunasExcel;
            List<string> colunas;
            if (string.IsNullOrWhiteSpace(json))
                colunas = new List<string> { "Nome", "Cargo", "Empresa" };
            else
                colunas = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(json);

            // Descobre o número de linhas preenchidas
            int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

            var pptApp = Globals.ThisAddIn.Application;
            var template = pptApp.ActivePresentation;

            var progressoForm = new ProgressoForm();
            progressoForm.Show();

            for (int i = 2; i <= lastRow; i++) // Começa na linha 2 (dados)
            {
                // Lê os valores das colunas para a linha atual
                var valores = new Dictionary<string, string>();
                for (int c = 0; c < colunas.Count; c++)
                {
                    var valorCelula = worksheet.Cells[i, c + 1].Text;
                    valores[colunas[c]] = valorCelula?.ToString() ?? string.Empty;
                }

                // Cria uma cópia do template
                string caminhoTemplate = template.FullName;
                var novaApresentacao = pptApp.Presentations.Open(caminhoTemplate, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                // Substitui palavras-chave em todos os slides
                foreach (PowerPoint.Slide slide in novaApresentacao.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                            shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextRange tr = shape.TextFrame.TextRange;

                            // Para cada coluna, substitui os padrões
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

                // Salva PowerPoint
                string nomeArquivoBase = valores[colunas[0]];
                string nomeArquivoPptx = System.IO.Path.Combine(pastaDestino, $"{PrimeiroUltimoNome(nomeArquivoBase).Replace(" ","")}.pptx");
                int contador = 0;
                while (System.IO.File.Exists(nomeArquivoPptx))
                {
                    contador++;
                    nomeArquivoPptx = System.IO.Path.Combine(pastaDestino, $"{PrimeiroUltimoNome(nomeArquivoBase).Replace(" ", "")} {contador}.pptx");
                }
                novaApresentacao.SaveAs(nomeArquivoPptx, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                // Salva PDF
                string nomeArquivoPdf = nomeArquivoPptx.Replace(".pptx",".pdf");
                novaApresentacao.SaveAs(nomeArquivoPdf, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);

                novaApresentacao.Close();

                progressoForm.AtualizarProgresso(i - 1, lastRow - 1);
            }

            workbook.Close(false);
            excelApp.Quit();

            progressoForm.ExibirConclusao();
        }

        private string PrimeiroUltimoNome(string nomeCompleto)
        {
            if (string.IsNullOrEmpty(nomeCompleto))
                return string.Empty;
            var partes = nomeCompleto.Trim().Split(' ');
            if (partes.Length == 0)
                return string.Empty;
            // Retorna o primeiro e o último nome
            return PrimeirasMaisculas($"{partes[0]} {partes[partes.Length - 1]}");
        }

        private string PrimeirasMaisculas(string texto)
        {
            if (string.IsNullOrEmpty(texto))
                return string.Empty;
            // Converte a primeira letra de cada palavra para maiúscula
            return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(texto.Trim().ToLower());
        }
        private void SubstituirTextoPreservandoFormatacao(PowerPoint.TextRange tr, string chave, string valor)
        {
            PowerPoint.TextRange encontrado = tr.Find(chave, 0, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            while (encontrado != null)
            {
                encontrado.Text = valor;
                encontrado = tr.Find(chave, 0, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            }
        }

        private void BtnParametros_Click(object sender, RibbonControlEventArgs e)
        {
            var frm = new ParametrosForm();
            frm.Show();
        }

        private void BtnEditarColunas_Click(object sender, RibbonControlEventArgs e)
        {
            using (var frm = new ColunasForm())
            {
                frm.ShowDialog();
            }
        }
    }
}
