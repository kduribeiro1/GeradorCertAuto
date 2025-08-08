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
    public partial class ParametrosForm : Form
    {
        public ParametrosForm(bool ehModeloSelecionado = false)
        {
            InitializeComponent();

            Properties.Settings.Default.Reload();

            if (ehModeloSelecionado)
            {
                var modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
                lblTituloFrm.Text = $"Parâmetros do Modelo: {modeloSelecionado.Nome}";
            }
            else
            {
                lblTituloFrm.Text = "Parâmetros do Excel";
            }

            this.Resize += ParametrosForm_Resize;

            // Configura o ListView
            listViewParametros.View = View.Details;
            listViewParametros.FullRowSelect = true;
            listViewParametros.Columns.Add("Parâmetro", 180);
            listViewParametros.Columns.Add("Descrição", 350);

            // Obtém os nomes das colunas da configuração
            var json = Properties.Settings.Default.ColunasExcel;
            if (ehModeloSelecionado)
            {
                var modeloSelecionado = JsonConvert.DeserializeObject<ModeloCertificado>(Properties.Settings.Default.ModeloSelecionado);
                json = modeloSelecionado.ColunasExcel;
            }

            List<string> colunas;
            if (string.IsNullOrWhiteSpace(json))
                colunas = new List<string> { "Nome", "Cargo", "Empresa" };
            else
                colunas = JsonConvert.DeserializeObject<List<string>>(json);

            // Adiciona os parâmetros e descrições dinamicamente
            foreach (var coluna in colunas)
            {
                listViewParametros.Items.Add(new ListViewItem(new[]
                {
                    $"<<{coluna}>>",
                    $"{coluna} conforme escrito no excel"
                }));
                listViewParametros.Items.Add(new ListViewItem(new[]
                {
                    $"<<{coluna}Maiusculo>>",
                    $"{coluna} em maiúsculas"
                }));
                listViewParametros.Items.Add(new ListViewItem(new[]
                {
                    $"<<{coluna}PriMaiusculo>>",
                    $"{coluna} com a primeira letra de cada palavra em maiúscula"
                }));
            }

            listViewParametros.Columns[1].Width = -2;
        }

        private void BtnCopiar_Click(object sender, EventArgs e)
        {
            if (listViewParametros.SelectedItems.Count > 0)
            {
                Clipboard.SetText(listViewParametros.SelectedItems[0].SubItems[0].Text);
            }
            else
            {
                MessageBox.Show("Selecione um parâmetro para copiar.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ParametrosForm_Resize(object sender, EventArgs e)
        {
            listViewParametros.Columns[1].Width = -2;
        }
    }
}
