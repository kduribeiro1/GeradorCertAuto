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
        public ParametrosForm()
        {
            InitializeComponent();

            this.Resize += ParametrosForm_Resize;

            // Configura o ListView
            listViewParametros.View = View.Details;
            listViewParametros.FullRowSelect = true;
            listViewParametros.Columns.Add("Parâmetro", 180);
            listViewParametros.Columns.Add("Descrição", 350);

            // Adiciona os parâmetros e descrições
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<Nome>>", "Nome completo do aluno conforme escrito no excel" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<NomeMaiusculo>>", "Nome completo do aluno em maiúsculas" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<NomePriMaiusculo>>", "Nome completo do aluno sendo com iniciais maiúsculas" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<Cargo>>", "Cargo do aluno conforme escrito no excel" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<CargoMaiusculo>>", "Cargo do aluno em maiúsculas" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<CargoPriMaiusculo>>", "Cargo do aluno sendo com iniciais maiúsculas" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<Empresa>>", "Empresa do aluno conforme escrito no excel" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<EmpresaMaiusculo>>", "Empresa do aluno em maiúsculas" }));
            listViewParametros.Items.Add(new ListViewItem(new[] { "<<EmpresaPriMaiusculo>>", "Empresa do aluno sendo com iniciais maiúsculas" }));
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
