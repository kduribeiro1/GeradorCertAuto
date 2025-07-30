using System;
using System.Windows.Forms;

namespace GeradorCertAuto
{
    public partial class ProgressoForm : Form
    {
        public ProgressoForm()
        {
            InitializeComponent();
        }

        public void AtualizarProgresso(int valor, int maximo)
        {
            progressBar.Value = valor;
            progressBar.Maximum = maximo;
            lblPercentual.Text = $"{(valor * 100 / (maximo == 0 ? 1 : maximo))}%";
            lblStatus.Text = $"Gerando {valor} de {maximo}...";
            Application.DoEvents();
        }

        public void ExibirConclusao()
        {
            lblConcluido.Text = "Geração dos certificados concluída!";
            lblConcluido.Visible = true;
            btnFechar.Visible = true;
            lblStatus.Visible = false;
            progressBar.Visible = false;
            lblPercentual.Visible = false;
        }

        private void btnFechar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}