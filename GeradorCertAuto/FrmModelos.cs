using Newtonsoft.Json;
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
    public partial class FrmModelos : Form
    {
        public FrmModelos()
        {
            InitializeComponent();
            CarregarModelos();

            btnEditar.Enabled = false;
            btnExcluir.Enabled = false;
            editarToolStripMenuItem.Enabled = false;
            excluirToolStripMenuItem.Enabled = false;
            btnBaixarModelo.Enabled = false;
            btnAtualizarModelo.Enabled = false;
            baixarModeloToolStripMenuItem.Enabled = false;
            atualizarModeloToolStripMenuItem.Enabled = false;


            listViewModelos.SelectedIndexChanged += ListViewModelos_SelectedIndexChanged;
            listViewModelos.MouseDown += ListViewModelos_MouseDown; // Adicione este evento
        }

        private void CarregarModelos()
        {
            listViewModelos.Items.Clear();
            var modelos = ConexaoBanco.GetModelos();
            foreach (var modelo in modelos)
            {
                var item = new ListViewItem(modelo.Id.ToString());
                item.SubItems.Add(modelo.Nome);
                item.SubItems.Add(modelo.Descricao);
                listViewModelos.Items.Add(item);
            }
        }

        private void BtnAtualizar_Click(object sender, EventArgs e)
        {
            CarregarModelos();
        }

        private void BtnNovo_Click(object sender, EventArgs e)
        {
            try
            {
                ModeloCertificado modelo = new ModeloCertificado();
                FrmModeloCert frm = new FrmModeloCert(modelo);
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    CarregarModelos();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao criar novo modelo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnEditar_Click(object sender, EventArgs e)
        {
            if (listViewModelos.SelectedItems.Count == 1)
            {
                int id = Convert.ToInt32(listViewModelos.SelectedItems[0].Text);
                ModeloCertificado modeloCertificado = new ModeloCertificado(id);
                modeloCertificado.GetModeloById();
                FrmModeloCert frm = new FrmModeloCert(modeloCertificado);
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    CarregarModelos();
                }
            }
        }

        private void BtnExcluir_Click(object sender, EventArgs e)
        {
            if (listViewModelos.SelectedItems.Count == 1)
            {
                int id = Convert.ToInt32(listViewModelos.SelectedItems[0].Text);
                var modelo = new ModeloCertificado { Id = id };
                if (MessageBox.Show("Deseja realmente excluir este modelo?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    modelo.DeleteModelo();
                    CarregarModelos();
                }
            }
        }

        private void BtnBaixarArquivo_Click(object sender, EventArgs e)
        {
            if (listViewModelos.SelectedItems.Count == 1)
            {
                int id = Convert.ToInt32(listViewModelos.SelectedItems[0].Text);
                var modelo = new ModeloCertificado { Id = id };
                modelo.GetModeloById(); // Carrega os dados do banco, incluindo o arquivo

                using (var sfd = new SaveFileDialog())
                {
                    string extensaoArquivo = System.IO.Path.GetExtension(modelo.NomeArquivo).ToLowerInvariant();
                    string filtro = string.Empty;

                    // Define o filtro conforme a extensão original
                    switch (extensaoArquivo)
                    {
                        case ".pptx":
                            filtro = "PowerPoint (*.pptx)|*.pptx";
                            break;
                        case ".ppt":
                            filtro = "PowerPoint 97-2003 (*.ppt)|*.ppt";
                            break;
                        case ".ppsx":
                            filtro = "PowerPoint Show (*.ppsx)|*.ppsx";
                            break;
                        case ".pps":
                            filtro = "PowerPoint Show 97-2003 (*.pps)|*.pps";
                            break;
                        default:
                            filtro = "Todos os arquivos (*.*)|*.*";
                            break;
                    }

                    sfd.DefaultExt = extensaoArquivo.TrimStart('.');
                    sfd.Filter = filtro;
                    sfd.FileName = $"{modelo.Nome}.{extensaoArquivo.TrimStart('.')}";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        modelo.SalvarArquivo(sfd.FileName);
                        MessageBox.Show("Arquivo salvo com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void BtnAlterarArquivo_Click(object sender, EventArgs e)
        {
            if (listViewModelos.SelectedItems.Count == 1)
            {
                int id = Convert.ToInt32(listViewModelos.SelectedItems[0].Text);
                var modelo = new ModeloCertificado { Id = id };
                modelo.GetModeloById(); // Carrega os dados atuais

                using (var ofd = new OpenFileDialog())
                {
                    ofd.Filter = "PowerPoint (*.pptx;*.ppt;*.ppsx;*.pps)|*.pptx;*.ppt;*.ppsx;*.pps";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        modelo.LerArquivo(ofd.FileName); // Atualiza a propriedade Arquivo
                        modelo.UpdateModelo(); // Atualiza no banco
                    }
                }
            }
        }

        private void ListViewModelos_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool selecionado = listViewModelos.SelectedItems.Count == 1;
            btnEditar.Enabled = selecionado;
            btnExcluir.Enabled = selecionado;
            editarToolStripMenuItem.Enabled = selecionado;
            excluirToolStripMenuItem.Enabled = selecionado;
            btnBaixarModelo.Enabled = selecionado;
            btnAtualizarModelo.Enabled = selecionado;
            baixarModeloToolStripMenuItem.Enabled = selecionado;
            atualizarModeloToolStripMenuItem.Enabled = selecionado;
        }

        private void ListViewModelos_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var item = listViewModelos.GetItemAt(e.X, e.Y);
                if (item != null)
                {
                    item.Selected = true;
                }
            }
        }
    }
}
