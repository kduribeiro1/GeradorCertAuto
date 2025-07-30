namespace GeradorCertAuto
{
    partial class ProgressoForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblPercentual;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblConcluido;
        private System.Windows.Forms.Button btnFechar;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblPercentual = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblConcluido = new System.Windows.Forms.Label();
            this.btnFechar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(20, 30);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(320, 23);
            this.progressBar.TabIndex = 0;
            // 
            // lblPercentual
            // 
            this.lblPercentual.AutoSize = true;
            this.lblPercentual.Location = new System.Drawing.Point(350, 35);
            this.lblPercentual.Name = "lblPercentual";
            this.lblPercentual.Size = new System.Drawing.Size(35, 13);
            this.lblPercentual.TabIndex = 1;
            this.lblPercentual.Text = "0%";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(20, 10);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 13);
            this.lblStatus.TabIndex = 2;
            // 
            // lblConcluido
            // 
            this.lblConcluido.AutoSize = true;
            this.lblConcluido.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.lblConcluido.ForeColor = System.Drawing.Color.Green;
            this.lblConcluido.Location = new System.Drawing.Point(20, 60);
            this.lblConcluido.Name = "lblConcluido";
            this.lblConcluido.Size = new System.Drawing.Size(0, 17);
            this.lblConcluido.TabIndex = 3;
            this.lblConcluido.Visible = false;
            // 
            // btnFechar
            // 
            this.btnFechar.Location = new System.Drawing.Point(320, 60);
            this.btnFechar.Name = "btnFechar";
            this.btnFechar.Size = new System.Drawing.Size(65, 25);
            this.btnFechar.TabIndex = 4;
            this.btnFechar.Text = "Fechar";
            this.btnFechar.UseVisualStyleBackColor = true;
            this.btnFechar.Visible = false;
            this.btnFechar.Click += new System.EventHandler(this.btnFechar_Click);
            // 
            // ProgressoForm
            // 
            this.ClientSize = new System.Drawing.Size(400, 100);
            this.Controls.Add(this.btnFechar);
            this.Controls.Add(this.lblConcluido);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblPercentual);
            this.Controls.Add(this.progressBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "ProgressoForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Progresso da Geração";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}