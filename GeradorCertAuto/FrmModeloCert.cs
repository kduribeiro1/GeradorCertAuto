using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GeradorCertAuto
{
    public partial class FrmModeloCert : Form
    {
        ModeloCertificado ModeloCertificado { get; set; }

        public FrmModeloCert()
        {
            InitializeComponent();
            Properties.Settings.Default.Reload();
            ModeloCertificado = new ModeloCertificado();
            lblTitulo.Text = "Novo Modelo de Certificado";
            btnSalvar.Text = "Cadastrar";
            txtNome.Clear();
            txtDescricao.Clear();
            txtArquivo.Clear();
            CarregarColunasConfig();
        }

        public FrmModeloCert(PowerPoint.Presentation modeloaberto)
        {
            InitializeComponent();

            Properties.Settings.Default.Reload();
            txtNome.Clear();
            txtDescricao.Clear();
            txtArquivo.Clear();
            var json = Properties.Settings.Default.ModeloSelecionado;
            if (!string.IsNullOrWhiteSpace(json))
            {
                ModeloCertificado = JsonConvert.DeserializeObject<ModeloCertificado>(json);
                ModeloCertificado.LerArquivo(modeloaberto);
                if (ModeloCertificado.Id.HasValue)
                {
                    lblTitulo.Text = $"Editar Modelo de Certificado: {ModeloCertificado.Id}";
                    btnSalvar.Text = "Atualizar";
                    txtNome.Text = ModeloCertificado.Nome;
                    txtDescricao.Text = ModeloCertificado.Descricao;
                    txtArquivo.Text = ModeloCertificado.Arquivo != null ? "Arquivo carregado" : "Nenhum arquivo carregado";
                    CarregarColunasModelo();
                }
                else
                {
                    lblTitulo.Text = "Novo Modelo de Certificado";
                    btnSalvar.Text = "Cadastrar";
                    txtArquivo.Clear();
                    txtArquivo.Text = ModeloCertificado.Arquivo != null ? "Arquivo carregado" : "Nenhum arquivo carregado";
                    CarregarColunasModelo();
                }
            }
            else
            {
                ModeloCertificado = new ModeloCertificado();
                ModeloCertificado.LerArquivo(modeloaberto);

                lblTitulo.Text = "Novo Modelo de Certificado";
                btnSalvar.Text = "Cadastrar";
                CarregarColunasModelo();
            }
        }

        public FrmModeloCert(ModeloCertificado modeloCertificado)
        {
            InitializeComponent();
            Properties.Settings.Default.Reload();
            if (modeloCertificado == null)
                throw new ArgumentNullException(nameof(modeloCertificado), "ModeloCertificado não pode ser nulo.");

            ModeloCertificado = modeloCertificado;

            txtNome.Clear();
            txtDescricao.Clear();
            txtArquivo.Clear();

            if (modeloCertificado.Id.HasValue)
            {
                lblTitulo.Text = $"Editar Modelo de Certificado: {modeloCertificado.Id}";
                btnSalvar.Text = "Atualizar";
                txtNome.Text = modeloCertificado.Nome;
                txtDescricao.Text = modeloCertificado.Descricao;
                txtArquivo.Text = modeloCertificado.Arquivo != null ? "Arquivo carregado" : "Nenhum arquivo carregado";
                CarregarColunasModelo();
            }
            else
            {
                lblTitulo.Text = "Novo Modelo de Certificado";
                btnSalvar.Text = "Cadastrar";
                CarregarColunasConfig();
            }
        }

        private void CarregarColunasConfig()
        {
            var json = Properties.Settings.Default.ColunasExcel;
            List<string> colunas;
            if (string.IsNullOrWhiteSpace(json))
                colunas = new List<string> { "Nome", "Cargo", "Empresa" };
            else
                colunas = JsonConvert.DeserializeObject<List<string>>(json);

            dgvColunas.Rows.Clear();
            foreach (var coluna in colunas)
                dgvColunas.Rows.Add(coluna);
        }

        private void CarregarColunasModelo()
        {
            var json = ModeloCertificado.ColunasExcel;
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

            dgvColunas.Rows.Clear();
            foreach (var coluna in colunas)
                dgvColunas.Rows.Add(coluna);
        }

        private void BtnSalvar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNome.Text))
            {
                MessageBox.Show("O nome do modelo é obrigatório.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (ModeloCertificado.Arquivo == null)
            {
                MessageBox.Show("Por favor, selecione um arquivo PowerPoint.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (dgvColunas.Rows.Count == 0)
            {
                MessageBox.Show("Por favor, adicione pelo menos uma coluna.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Deseja salvar as alterações?", "Salvar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var colunas = new List<string>();
                foreach (DataGridViewRow row in dgvColunas.Rows)
                {
                    if (row.Cells[0].Value != null && !string.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                        colunas.Add(row.Cells[0].Value.ToString());
                }
                ModeloCertificado.ColunasExcel = JsonConvert.SerializeObject(colunas);
                ModeloCertificado.Nome = txtNome.Text.Trim();
                ModeloCertificado.Descricao = txtDescricao.Text?.Trim();

                if (ModeloCertificado.Id.HasValue)
                {
                    ModeloCertificado.UpdateModelo();

                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    ModeloCertificado.InserirNovoModelo();

                    DialogResult = DialogResult.OK;
                    Close();
                }
            }
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja cancelar as alterações?", "Cancelar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void BtnSelecionar_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrWhiteSpace(openFileDialog1.FileName))
                {
                    MessageBox.Show("Nenhum arquivo selecionado.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string ext = System.IO.Path.GetExtension(openFileDialog1.FileName).ToLowerInvariant();
                if (ext != ".pptx" && ext != ".ppt" && ext != ".ppsx" && ext != ".pps")
                {
                    MessageBox.Show("Por favor, selecione um arquivo PowerPoint válido.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (File.Exists(openFileDialog1.FileName))
                {
                    txtArquivo.Text = openFileDialog1.FileName;
                    ModeloCertificado.LerArquivo(openFileDialog1.FileName);
                }
                else
                {
                    MessageBox.Show("Arquivo não encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtArquivo.Clear();
                }
            }
        }

        private void BtnAdicionar_Click(object sender, EventArgs e)
        {
            dgvColunas.Rows.Add("");
        }

        private void BtnRemover_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvColunas.SelectedRows)
            {
                if (!row.IsNewRow)
                    dgvColunas.Rows.Remove(row);
            }
        }
    }
}
