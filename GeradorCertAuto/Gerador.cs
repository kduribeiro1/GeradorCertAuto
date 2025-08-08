using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GeradorCertAuto
{
    public partial class Gerador
    {
        private PowerPoint.Presentation CertificadoSelecionado = null;

        private void Gerador_Load(object sender, RibbonUIEventArgs e)
        {
            Properties.Settings.Default.Reload();
            AtualizarListaModelos();
        }

        private void AtualizarListaModelos()
        {
            Properties.Settings.Default.Reload();
            dpdModelos.Items.Clear();
            var modelos = ConexaoBanco.GetModelos();

            var item0 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item0.Label = "Nenhum modelo selecionado";
            item0.Tag = 0; // ou qualquer identificador
            dpdModelos.Items.Add(item0);
            foreach (var modelo in modelos)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = modelo.Nome;
                item.Tag = modelo.Id; // ou qualquer identificador
                dpdModelos.Items.Add(item);
            }
            if (Properties.Settings.Default.ModeloSelecionado != string.Empty)
            {
                // Deserializa o modelo selecionado
                var modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
                if (modeloSelecionado != null)
                {
                    var itemSelecionado = dpdModelos.Items.FirstOrDefault(i => (int?)i.Tag == modeloSelecionado.Id);
                    if (itemSelecionado != null)
                    {
                        dpdModelos.SelectedItem = itemSelecionado;
                    }
                    else
                    {
                        dpdModelos.SelectedItem = dpdModelos.Items[0];
                    }
                }
            }
            else
            {
                if (dpdModelos.Items.Count > 0)
                {
                    dpdModelos.SelectedItem = dpdModelos.Items[0];
                }
            }
        }

        private void DpdModelos_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var selectedItem = dpdModelos.SelectedItem;
            if (selectedItem != null)
            {
                int modeloId = (int)selectedItem.Tag;
                // Busca o modelo completo no banco
                var modelo = new ModeloCertificado(modeloId);
                modelo.GetModeloById();

                // Serializa para JSON
                string json = JsonConvert.SerializeObject(modelo);

                // Salva na configuração
                Properties.Settings.Default.ModeloSelecionado = json;
                Properties.Settings.Default.Save(); 

                CertificadoSelecionado = modelo.AbrirArquivo();
                if (CertificadoSelecionado != null)
                {
                    // Garante que a janela da apresentação fique em foco
                    if (CertificadoSelecionado.Windows.Count > 0)
                    {
                        CertificadoSelecionado.Windows[1].Activate();
                    }
                }
                else
                {
                    MessageBox.Show("Não foi possível abrir o modelo selecionado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                CertificadoSelecionado = null;
                Properties.Settings.Default.ModeloSelecionado = string.Empty;
                Properties.Settings.Default.Save();
            }
            AtualizarListaModelos();
        }

        private void BtnGerarExcel_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();
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
            Properties.Settings.Default.Reload();
            var pptApp = Globals.ThisAddIn.Application;
            var presentation = pptApp.ActivePresentation;

            // Verifica se a apresentação foi salva após modificações
            if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                MessageBox.Show("A apresentação precisa ser salva antes de gerar os certificados.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string pastaRaiz = System.IO.Path.GetDirectoryName(presentation.FullName);
            string pasta = System.IO.Path.Combine(pastaRaiz, "Certificados");
            if (!System.IO.Directory.Exists(pasta))
            {
                System.IO.Directory.CreateDirectory(pasta);
            }
            string caminhoExcel = System.IO.Path.Combine(pastaRaiz, "TabelaAlunos.xlsx");
            if (!File.Exists(caminhoExcel))
            {
                MessageBox.Show("O arquivo TabelaAlunos.xlsx não foi encontrado. Por favor, gere a tabela primeiro.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
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

        private void BtnAbrirSistema_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();
            using (var frm = new FrmSistema())
            {
                frm.ShowDialog();
            }
        }

        private void BtnModelosCads_Click(object sender, RibbonControlEventArgs e)
        {
            using (FrmModelos frm = new FrmModelos())
            {
                frm.ShowDialog();
                AtualizarListaModelos();
            }
        }

        private void BtnCertSelectEdtCol_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();
            var modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
            if (modeloSelecionado == null)
            {
                MessageBox.Show("Nenhum modelo selecionado. Por favor, selecione um modelo primeiro.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var frm = new ColunasForm(true))
            {
                frm.ShowDialog();
            }
        }

        private void BtnCertSelectGerTab_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();
            var modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
            if (modeloSelecionado == null)
            {
                MessageBox.Show("Nenhum modelo selecionado. Por favor, selecione um modelo primeiro.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBox.Show($"Deseja gerar uma tabela de alunos usando o modelo selecionado ({modeloSelecionado.Nome})?", "Gerar Tabela", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            var workbook = excelApp.Workbooks.Add();
            var worksheet = (Excel.Worksheet)workbook.Sheets[1];
            worksheet.Name = modeloSelecionado.Nome.Replace(" ","_").Substring(0,15);

            // Obtém os nomes das colunas da configuração

            var json = modeloSelecionado.ColunasExcel;
            List<string> colunas;
            if (string.IsNullOrWhiteSpace(json))
            {
                json = Properties.Settings.Default.ColunasExcel;
                if (string.IsNullOrWhiteSpace(json))
                {
                    colunas = new List<string> { "Nome", "Cargo", "Empresa" };
                }
                else
                {
                    colunas = JsonConvert.DeserializeObject<List<string>>(json);
                }
            }
            else
            {
                colunas = JsonConvert.DeserializeObject<List<string>>(json);
            }

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

        private void BtnCertSelectShowPar_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();
            var modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
            if (modeloSelecionado == null)
            {
                MessageBox.Show("Nenhum modelo selecionado. Por favor, selecione um modelo primeiro.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var frm = new ParametrosForm(true);
            frm.Show();
        }

        private void BtnCertSelectGerCert_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();

            // Tenta obter o modelo selecionado
            ModeloCertificado modeloSelecionado = null;
            if (!string.IsNullOrWhiteSpace(Properties.Settings.Default.ModeloSelecionado))
            {
                modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
            }

            if (modeloSelecionado == null || modeloSelecionado.Id == null)
            {
                MessageBox.Show("Nenhum modelo selecionado. Por favor, selecione um modelo primeiro.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBox.Show($"Deseja gerar certificados usando o modelo selecionado ({modeloSelecionado.Nome})?", "Gerar Certificados", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            var pptApp = Globals.ThisAddIn.Application;
            var presentation = pptApp.ActivePresentation;

            // Verifica se a apresentação foi salva após modificações
            if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                MessageBox.Show("A apresentação precisa ser salva antes de gerar os certificados.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string pastaRaiz = System.IO.Path.GetDirectoryName(presentation.FullName);
            string pasta = System.IO.Path.Combine(pastaRaiz, "Certificados");
            if (!System.IO.Directory.Exists(pasta))
            {
                System.IO.Directory.CreateDirectory(pasta);
            }
            string caminhoExcel = System.IO.Path.Combine(pastaRaiz, "TabelaAlunos.xlsx");

            if (!File.Exists(caminhoExcel))
            {
                MessageBox.Show("O arquivo TabelaAlunos.xlsx não foi encontrado. Por favor, gere a tabela primeiro.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

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
                string nomeArquivoPptx = System.IO.Path.Combine(pasta, $"{PrimeiroUltimoNome(nomeArquivoBase).Replace(" ", "")}.pptx");
                int contador = 0;
                while (System.IO.File.Exists(nomeArquivoPptx))
                {
                    contador++;
                    nomeArquivoPptx = System.IO.Path.Combine(pasta, $"{PrimeiroUltimoNome(nomeArquivoBase).Replace(" ", "")} {contador}.pptx");
                }
                novaApresentacao.SaveAs(nomeArquivoPptx, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                // Salva PDF
                string nomeArquivoPdf = nomeArquivoPptx.Replace(".pptx", ".pdf");
                novaApresentacao.SaveAs(nomeArquivoPdf, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);

                novaApresentacao.Close();

                progressoForm.AtualizarProgresso(i - 1, lastRow - 1);
            }

            workbook.Close(false);
            excelApp.Quit();

            progressoForm.ExibirConclusao();
        }

        private void BtnCertSelectSalvar_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.Reload();
            var pptApp = Globals.ThisAddIn.Application;
            var presentation = pptApp.ActivePresentation;

            // Verifica se a apresentação foi salva após modificações
            if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                MessageBox.Show("A apresentação precisa ser salva antes de gerar os certificados.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }



            // Tenta obter o modelo selecionado
            ModeloCertificado modeloSelecionado = null;
            if (!string.IsNullOrWhiteSpace(Properties.Settings.Default.ModeloSelecionado))
            {
                modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
            }

            if (modeloSelecionado != null && modeloSelecionado.Id.HasValue)
            {
                if (MessageBox.Show($"Deseja atualizar o modelo selecionado ({modeloSelecionado.Nome}) com a apresentação atual?", "Atualizar Modelo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
                // Atualiza o arquivo do modelo com a apresentação atual
                modeloSelecionado.LerArquivo(presentation.FullName);
                modeloSelecionado.UpdateModelo();
                MessageBox.Show("Modelo atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // Cria um novo modelo preenchendo o arquivo com a apresentação atual
                var novoModelo = new ModeloCertificado();
                novoModelo.LerArquivo(presentation.FullName);

                using (var frm = new FrmModeloCert(novoModelo))
                {
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        AtualizarListaModelos();
                    }
                }
            }

        }
    }
}
