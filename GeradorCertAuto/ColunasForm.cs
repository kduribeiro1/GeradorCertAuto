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
    public partial class ColunasForm : Form
    {
        public ColunasForm()
        {
            InitializeComponent();
            CarregarColunas();
        }

        private void CarregarColunas()
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

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            dgvColunas.Rows.Add("");
        }

        private void btnRemover_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvColunas.SelectedRows)
            {
                if (!row.IsNewRow)
                    dgvColunas.Rows.Remove(row);
            }
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            var colunas = new List<string>();
            foreach (DataGridViewRow row in dgvColunas.Rows)
            {
                if (row.Cells[0].Value != null && !string.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                    colunas.Add(row.Cells[0].Value.ToString());
            }
            Properties.Settings.Default.ColunasExcel = JsonConvert.SerializeObject(colunas);
            Properties.Settings.Default.Save();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
