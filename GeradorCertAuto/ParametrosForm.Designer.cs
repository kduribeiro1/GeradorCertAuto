namespace GeradorCertAuto
{
    partial class ParametrosForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnCopiar = new System.Windows.Forms.Button();
            this.listViewParametros = new System.Windows.Forms.ListView();
            this.lblTituloFrm = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.Controls.Add(this.btnCopiar, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.listViewParametros, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.lblTituloFrm, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 7;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(559, 390);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // btnCopiar
            // 
            this.btnCopiar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnCopiar.Location = new System.Drawing.Point(8, 353);
            this.btnCopiar.Name = "btnCopiar";
            this.btnCopiar.Size = new System.Drawing.Size(543, 29);
            this.btnCopiar.TabIndex = 0;
            this.btnCopiar.Text = "Copiar Parâmetro Selecionado";
            this.btnCopiar.UseVisualStyleBackColor = true;
            this.btnCopiar.Click += new System.EventHandler(this.BtnCopiar_Click);
            // 
            // listViewParametros
            // 
            this.listViewParametros.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewParametros.HideSelection = false;
            this.listViewParametros.Location = new System.Drawing.Point(8, 48);
            this.listViewParametros.Name = "listViewParametros";
            this.listViewParametros.Size = new System.Drawing.Size(543, 294);
            this.listViewParametros.TabIndex = 1;
            this.listViewParametros.UseCompatibleStateImageBehavior = false;
            // 
            // lblTituloFrm
            // 
            this.lblTituloFrm.AutoSize = true;
            this.lblTituloFrm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTituloFrm.Location = new System.Drawing.Point(8, 5);
            this.lblTituloFrm.Name = "lblTituloFrm";
            this.lblTituloFrm.Size = new System.Drawing.Size(543, 35);
            this.lblTituloFrm.TabIndex = 2;
            this.lblTituloFrm.Text = "Parâmetros do Excel";
            this.lblTituloFrm.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ParametrosForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(559, 390);
            this.Controls.Add(this.tableLayoutPanel1);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(441, 402);
            this.Name = "ParametrosForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Parâmetros";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btnCopiar;
        private System.Windows.Forms.ListView listViewParametros;
        private System.Windows.Forms.Label lblTituloFrm;
    }
}