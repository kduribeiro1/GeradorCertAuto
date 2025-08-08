using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows.Forms;

namespace GeradorCertAuto
{
    internal static class ConexaoBanco
    {
        // Caminho do banco (pode ser relativo ou absoluto)
        private static string pastaApp = AppDomain.CurrentDomain.BaseDirectory;
        private static string caminhoBanco = System.IO.Path.Combine(pastaApp, "certificados.db");
        private static string DbPath = $"Data Source={caminhoBanco};Version=3;";

        internal static void InserirNovoModelo(this ModeloCertificado modeloCertificado)
        {
            try
            {
                bool existeNome = false;
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = "SELECT COUNT(*) FROM ModelosCertificado WHERE Nome=@nome";
                        cmd.Parameters.Add(modeloCertificado.GetParameter(nameof(ModeloCertificado.Nome)));
                        int count = Convert.ToInt32(cmd.ExecuteScalar());
                        if (count > 0)
                        {
                            existeNome = true;
                        }
                    }
                }
                if (existeNome)
                {
                    MessageBox.Show("Já existe um modelo com esse nome. Por favor, escolha outro nome.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = modeloCertificado.GetScriptInsert();
                        cmd.Parameters.AddRange(modeloCertificado.GetParametersInsert());
                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Novo modelo inserido com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = "SELECT Id FROM ModelosCertificado WHERE Nome=@nome";
                        cmd.Parameters.Add(modeloCertificado.GetParameter(nameof(ModeloCertificado.Nome)));
                        modeloCertificado.Id = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao inserir novo modelo: {ex.Message}", "Falha", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal static void UpdateModelo(this ModeloCertificado modeloCertificado)
        {
            try
            {
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = "SELECT COUNT(*) FROM ModelosCertificado WHERE Nome=@nome AND Id!=@id";
                        cmd.Parameters.Add(modeloCertificado.GetParameter(nameof(ModeloCertificado.Nome)));
                        cmd.Parameters.Add(modeloCertificado.GetParameter(nameof(ModeloCertificado.Id)));
                        int count = Convert.ToInt32(cmd.ExecuteScalar());
                        if (count > 0)
                        {
                            MessageBox.Show("Já existe um modelo com esse nome. Por favor, escolha outro nome.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }

                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = modeloCertificado.GetScriptUpdate();
                        cmd.Parameters.AddRange(modeloCertificado.GetParametersUpdate());
                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("O modelo foi atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar o modelo: {ex.Message}", "Falha", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal static void DeleteModelo(this ModeloCertificado modeloCertificado)
        {
            try
            {
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = modeloCertificado.GetScriptDelete();
                        cmd.Parameters.AddRange(modeloCertificado.GetParametersDelete());
                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("O modelo foi excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao excluir o modelo: {ex.Message}", "Falha", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal static void GetModeloById(this ModeloCertificado modeloCertificado)
        {
            try
            {
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = modeloCertificado.GetScriptSelect();
                        cmd.Parameters.AddRange(modeloCertificado.GetParametersSelect());
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                modeloCertificado.Nome = reader.GetString(1);
                                modeloCertificado.Descricao = reader.IsDBNull(2) ? null : reader.GetString(2);
                                modeloCertificado.ColunasExcel = reader.GetString(3);
                                modeloCertificado.NomeArquivo = reader.IsDBNull(4) ? null : reader.GetString(4);
                                modeloCertificado.Arquivo = reader.IsDBNull(5) ? null : (byte[])reader[5];
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao buscar modelo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal static List<ModeloCertificado> GetModelos()
        {
            var modelos = new List<ModeloCertificado>();
            try
            {
                using (var conn = new SQLiteConnection(DbPath))
                {
                    using (var cmd = new SQLiteCommand(conn))
                    {
                        conn.Open();
                        cmd.CommandText = ModeloCertificado.GetScriptSelectAll();
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var modelo = new ModeloCertificado
                                {
                                    Id = reader.GetInt32(0),
                                    Nome = reader.GetString(1),
                                    Descricao = reader.IsDBNull(2) ? null : reader.GetString(2),
                                    ColunasExcel = reader.GetString(3),
                                    NomeArquivo = reader.IsDBNull(4) ? null : reader.GetString(4),
                                    Arquivo = reader.IsDBNull(5) ? null : (byte[])reader[5]
                                };
                                modelos.Add(modelo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao buscar modelos: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return modelos;
        }
    }
}
