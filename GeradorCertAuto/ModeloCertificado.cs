using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GeradorCertAuto
{
    public class ModeloCertificado
    {
        public int? Id { get; set; }
        public string Nome { get; set; }
        public string Descricao { get; set; }
        public string ColunasExcel { get; set; }
        public string NomeArquivo { get; set; }
        public byte[] Arquivo { get; set; }

        public ModeloCertificado()
        {
            Id = null;
            Nome = string.Empty;
            Descricao = string.Empty;
            ColunasExcel = string.Empty;
            NomeArquivo = string.Empty;
            Arquivo = null;
        }

        public ModeloCertificado(int id)
        {
            Id = id;
            Nome = string.Empty;
            Descricao = string.Empty;
            ColunasExcel = string.Empty;
            NomeArquivo = string.Empty;
            Arquivo = null;
        }

        public ModeloCertificado(string nome, string descricao, string colunasExcel)
        {
            Id = null;
            Nome = nome;
            Descricao = descricao;
            ColunasExcel = colunasExcel;
            NomeArquivo = string.Empty;
            Arquivo = null;
        }

        public ModeloCertificado(string nome, string descricao, string colunasExcel, string nomeArquivo, byte[] arquivo)
        {
            Id = null;
            Nome = nome;
            Descricao = descricao;
            ColunasExcel = colunasExcel;
            NomeArquivo = nomeArquivo;
            Arquivo = arquivo;
        }

        public ModeloCertificado(int id, string nome, string descricao, string colunasExcel, string nomeArquivo, byte[] arquivo)
        {
            Id = id;
            Nome = nome;
            Descricao = descricao;
            ColunasExcel = colunasExcel;
            NomeArquivo = nomeArquivo;
            Arquivo = arquivo;
        }

        public void LerArquivo(string caminhoArquivo)
        {
            if (string.IsNullOrEmpty(caminhoArquivo))
                throw new ArgumentException("Caminho do arquivo não pode ser nulo ou vazio.", nameof(caminhoArquivo));
            try
            {
                NomeArquivo = System.IO.Path.GetFileName(caminhoArquivo);
                Arquivo = System.IO.File.ReadAllBytes(caminhoArquivo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao ler o arquivo: {ex.Message}", ex);
            }
        }

        public void LerArquivo(PowerPoint.Presentation arquivo)
        {
            if (arquivo == null)
                throw new ArgumentException("Arquivo não reconhecido.", nameof(arquivo));
            try
            {
                // Verifica se a apresentação foi salva após modificações
                if (arquivo.Saved == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    MessageBox.Show("A apresentação precisa ser salva antes de ser carregada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Pega o caminho do arquivo pptx salvo
                string caminho = arquivo.FullName;
                NomeArquivo = System.IO.Path.GetFileName(caminho);
                Arquivo = System.IO.File.ReadAllBytes(caminho);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao ler o arquivo: {ex.Message}", ex);
            }
        }

        public void SalvarArquivo(string caminhoArquivo = "")
        {
            if (string.IsNullOrEmpty(caminhoArquivo))
            {
                caminhoArquivo = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Modelos", NomeArquivo);
            }

            if (Arquivo == null || Arquivo.Length == 0)
                throw new InvalidOperationException("Nenhum arquivo carregado para salvar.");
            try
            {
                string diretorio = System.IO.Path.GetDirectoryName(caminhoArquivo);
                if (!System.IO.Directory.Exists(diretorio))
                {
                    System.IO.Directory.CreateDirectory(diretorio);
                }
                string nomeArquivo = System.IO.Path.GetFileName(caminhoArquivo);
                string extensao = System.IO.Path.GetExtension(nomeArquivo).ToLowerInvariant();
                string extensaoArquivo = System.IO.Path.GetExtension(NomeArquivo).ToLowerInvariant();
                if (extensao != extensaoArquivo)
                {
                    nomeArquivo = System.IO.Path.ChangeExtension(nomeArquivo, extensaoArquivo);
                }
                if (System.IO.File.Exists(System.IO.Path.Combine(diretorio, nomeArquivo)))
                {
                    System.IO.File.Delete(System.IO.Path.Combine(diretorio, nomeArquivo));
                }
                System.IO.File.WriteAllBytes(System.IO.Path.Combine(diretorio, nomeArquivo), Arquivo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar o arquivo: {ex.Message}", ex);
            }
        }

        public PowerPoint.Presentation AbrirArquivo(string caminhoArquivo = "")
        {

            if (string.IsNullOrEmpty(caminhoArquivo))
            {
                caminhoArquivo = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Modelos", NomeArquivo);
            }    

            if (Arquivo == null || Arquivo.Length == 0)
                MessageBox.Show("Nenhum arquivo carregado para salvar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            try
            {
                string diretorio = System.IO.Path.GetDirectoryName(caminhoArquivo);
                if (!System.IO.Directory.Exists(diretorio))
                {
                    System.IO.Directory.CreateDirectory(diretorio);
                }
                string nomeArquivo = System.IO.Path.GetFileName(caminhoArquivo);
                string extensao = System.IO.Path.GetExtension(nomeArquivo).ToLowerInvariant();
                string extensaoArquivo = System.IO.Path.GetExtension(NomeArquivo).ToLowerInvariant();
                if (extensao != extensaoArquivo)
                {
                    nomeArquivo = System.IO.Path.ChangeExtension(nomeArquivo, extensaoArquivo);
                }
                string caminhoTemplate = System.IO.Path.Combine(diretorio, nomeArquivo);
                if (System.IO.File.Exists(caminhoTemplate))
                {
                    System.IO.File.Delete(caminhoTemplate);
                }
                System.IO.File.WriteAllBytes(caminhoTemplate, Arquivo);
                var pptApp = Globals.ThisAddIn.Application;
                return pptApp.Presentations.Open(caminhoTemplate, WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar o arquivo: {ex.Message}", ex);
            }
        }

        public string GetScriptInsert()
        {
            return "INSERT INTO ModelosCertificado (Nome, Descricao, ColunasExcel, NomeArquivo, ArquivoModelo) VALUES (@nome, @descricao, @colunasexcel, @nomearquivo, @arquivo)";
        }

        public SQLiteParameter[] GetParametersInsert()
        {
            return new SQLiteParameter[]
            {
                new SQLiteParameter("@nome", Nome),
                new SQLiteParameter("@descricao", Descricao ?? (object)DBNull.Value),
                new SQLiteParameter("@colunasexcel", ColunasExcel),
                new SQLiteParameter("@arquivo", Arquivo),
                new SQLiteParameter("@nomearquivo", NomeArquivo)
            };
        }

        public string GetScriptUpdate()
        {
            return "UPDATE ModelosCertificado SET Nome=@nome, Descricao=@descricao, ColunasExcel=@colunasexcel, NomeArquivo=@nomearquivo, ArquivoModelo=@arquivo WHERE Id=@id";
        }

        public SQLiteParameter[] GetParametersUpdate()
        {
            return new SQLiteParameter[]
            {
                new SQLiteParameter("@id", Id),
                new SQLiteParameter("@nome", Nome),
                new SQLiteParameter("@descricao", Descricao ?? (object)DBNull.Value),
                new SQLiteParameter("@colunasexcel", ColunasExcel),
                new SQLiteParameter("@arquivo", Arquivo),
                new SQLiteParameter("@nomearquivo", NomeArquivo)
            };
        }

        public string GetScriptDelete()
        {
            return "DELETE FROM ModelosCertificado WHERE Id=@id";
        }

        public SQLiteParameter[] GetParametersDelete()
        {
            return new SQLiteParameter[]
            {
                new SQLiteParameter("@id", Id)
            };
        }

        public string GetScriptSelect()
        {
            return "SELECT Id, Nome, Descricao, ColunasExcel, NomeArquivo, ArquivoModelo FROM ModelosCertificado WHERE Id=@id";
        }

        public SQLiteParameter[] GetParametersSelect()
        {
            return new SQLiteParameter[]
            {
                new SQLiteParameter("@id", Id)
            };
        }

        public static string GetScriptSelectAll()
        {
            return "SELECT Id, Nome, Descricao, ColunasExcel, NomeArquivo, ArquivoModelo FROM ModelosCertificado";
        }

        public SQLiteParameter GetParameter(string titulocoluna)
        {
            switch (titulocoluna)
            {
                case nameof(Id):
                    return new SQLiteParameter("@id", Id);
                case nameof(Nome):
                    return new SQLiteParameter("@nome", Nome);
                case nameof(Descricao):
                    return new SQLiteParameter("@descricao", Descricao ?? (object)DBNull.Value);
                case nameof(ColunasExcel):
                    return new SQLiteParameter("@colunasexcel", ColunasExcel);
                case nameof(Arquivo):
                    return new SQLiteParameter("@arquivo", Arquivo);
                case nameof(NomeArquivo):
                    return new SQLiteParameter("@nomearquivo", NomeArquivo);
                default:
                    throw new ArgumentException($"Coluna desconhecida: {titulocoluna}", nameof(titulocoluna));
            }
        }

        public override string ToString()
        {
            return $"{Id} - {Nome}";
        }
    }
}
