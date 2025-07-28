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

            // Cabeçalhos da tabela
            worksheet.Cells[1, 1] = "Nome";
            worksheet.Cells[1, 2] = "Cargo";
            worksheet.Cells[1, 3] = "Empresa";

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

            // Descobre o número de linhas preenchidas
            int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

            var pptApp = Globals.ThisAddIn.Application;
            var template = pptApp.ActivePresentation;

            for (int i = 2; i <= lastRow; i++) // Começa na linha 2 (dados)
            {
                string nome = worksheet.Cells[i, 1].Text;
                string cargo = worksheet.Cells[i, 2].Text;
                string empresa = worksheet.Cells[i, 3].Text;

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
                            if (shape.TextFrame.TextRange.Text.Contains("<<Nome>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<Cargo>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<Empresa>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<NomeMaiusculo>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<NomePriMaiusculo>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<CargoMaiusculo>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<CargoPriMaiusculo>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<EmpresaMaiusculo>>") ||
                                shape.TextFrame.TextRange.Text.Contains("<<EmpresaPriMaiusculo>>")
                                )
                            {
                                PowerPoint.TextRange tr = shape.TextFrame.TextRange;

                                // Substitui cada parâmetro individualmente
                                SubstituirTextoPreservandoFormatacao(tr, "<<NomeMaiusculo>>", nome.Trim().ToUpper());
                                SubstituirTextoPreservandoFormatacao(tr, "<<NomePriMaiusculo>>", PrimeirasMaisculas(nome));
                                SubstituirTextoPreservandoFormatacao(tr, "<<Nome>>", nome.Trim());
                                SubstituirTextoPreservandoFormatacao(tr, "<<CargoMaiusculo>>", cargo.Trim().ToUpper());
                                SubstituirTextoPreservandoFormatacao(tr, "<<CargoPriMaiusculo>>", PrimeirasMaisculas(cargo));
                                SubstituirTextoPreservandoFormatacao(tr, "<<Cargo>>", cargo.Trim());
                                SubstituirTextoPreservandoFormatacao(tr, "<<EmpresaMaiusculo>>", empresa.Trim().ToUpper());
                                SubstituirTextoPreservandoFormatacao(tr, "<<EmpresaPriMaiusculo>>", PrimeirasMaisculas(empresa));
                                SubstituirTextoPreservandoFormatacao(tr, "<<Empresa>>", empresa.Trim());
                            }
                        }
                    }
                }

                // Salva PowerPoint
                string nomeArquivoPptx = System.IO.Path.Combine(pastaDestino, $"{PrimeiroUltimoNome(nome).Replace(" ","")}.pptx");
                int contador = 0;
                while (System.IO.File.Exists(nomeArquivoPptx))
                {
                    contador++;
                    nomeArquivoPptx = System.IO.Path.Combine(pastaDestino, $"{PrimeiroUltimoNome(nome).Replace(" ", "")} {contador}.pptx");
                }
                novaApresentacao.SaveAs(nomeArquivoPptx, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                // Salva PDF
                string nomeArquivoPdf = nomeArquivoPptx.Replace(".pptx",".pdf");
                novaApresentacao.SaveAs(nomeArquivoPdf, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);

                novaApresentacao.Close();
            }

            workbook.Close(false);
            excelApp.Quit();

            System.Windows.Forms.MessageBox.Show("Certificados foram gerados na pasta: " + pastaDestino);
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
    }
}
