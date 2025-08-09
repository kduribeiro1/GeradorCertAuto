namespace GeradorCertAuto
{
    partial class FrmModelos
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ListView listViewModelos;
        private System.Windows.Forms.ColumnHeader columnHeaderId;
        private System.Windows.Forms.ColumnHeader columnHeaderNome;
        private System.Windows.Forms.ColumnHeader columnHeaderDescricao;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.listViewModelos = new System.Windows.Forms.ListView();
            this.columnHeaderId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderNome = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderColunas = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderDescricao = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.novoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.baixarModeloToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.atualizarModeloToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excluirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.atualizarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.btnNovo = new System.Windows.Forms.ToolStripMenuItem();
            this.btnEditar = new System.Windows.Forms.ToolStripMenuItem();
            this.btnBaixarModelo = new System.Windows.Forms.ToolStripMenuItem();
            this.btnAtualizarModelo = new System.Windows.Forms.ToolStripMenuItem();
            this.btnExcluir = new System.Windows.Forms.ToolStripMenuItem();
            this.btnAtualizar = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listViewModelos
            // 
            this.listViewModelos.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderId,
            this.columnHeaderNome,
            this.columnHeaderColunas,
            this.columnHeaderDescricao});
            this.listViewModelos.ContextMenuStrip = this.contextMenuStrip1;
            this.listViewModelos.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewModelos.FullRowSelect = true;
            this.listViewModelos.GridLines = true;
            this.listViewModelos.HideSelection = false;
            this.listViewModelos.Location = new System.Drawing.Point(8, 8);
            this.listViewModelos.MultiSelect = false;
            this.listViewModelos.Name = "listViewModelos";
            this.listViewModelos.Size = new System.Drawing.Size(612, 312);
            this.listViewModelos.TabIndex = 0;
            this.listViewModelos.UseCompatibleStateImageBehavior = false;
            this.listViewModelos.View = System.Windows.Forms.View.Details;
            // 
            // columnHeaderId
            // 
            this.columnHeaderId.Text = "Id";
            this.columnHeaderId.Width = 50;
            // 
            // columnHeaderNome
            // 
            this.columnHeaderNome.Text = "Nome";
            this.columnHeaderNome.Width = 190;
            // 
            // columnHeaderColunas
            // 
            this.columnHeaderColunas.Text = "Colunas Excel";
            this.columnHeaderColunas.Width = 200;
            // 
            // columnHeaderDescricao
            // 
            this.columnHeaderDescricao.Text = "Descrição";
            this.columnHeaderDescricao.Width = 300;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.novoToolStripMenuItem,
            this.editarToolStripMenuItem,
            this.baixarModeloToolStripMenuItem,
            this.atualizarModeloToolStripMenuItem,
            this.excluirToolStripMenuItem,
            this.atualizarToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(165, 136);
            // 
            // novoToolStripMenuItem
            // 
            this.novoToolStripMenuItem.Name = "novoToolStripMenuItem";
            this.novoToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.novoToolStripMenuItem.Text = "Novo";
            this.novoToolStripMenuItem.Click += new System.EventHandler(this.BtnNovo_Click);
            // 
            // editarToolStripMenuItem
            // 
            this.editarToolStripMenuItem.Name = "editarToolStripMenuItem";
            this.editarToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.editarToolStripMenuItem.Text = "Editar";
            this.editarToolStripMenuItem.Click += new System.EventHandler(this.BtnEditar_Click);
            // 
            // baixarModeloToolStripMenuItem
            // 
            this.baixarModeloToolStripMenuItem.Name = "baixarModeloToolStripMenuItem";
            this.baixarModeloToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.baixarModeloToolStripMenuItem.Text = "Baixar Modelo";
            this.baixarModeloToolStripMenuItem.Click += new System.EventHandler(this.BtnBaixarArquivo_Click);
            // 
            // atualizarModeloToolStripMenuItem
            // 
            this.atualizarModeloToolStripMenuItem.Name = "atualizarModeloToolStripMenuItem";
            this.atualizarModeloToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.atualizarModeloToolStripMenuItem.Text = "Atualizar Modelo";
            this.atualizarModeloToolStripMenuItem.Click += new System.EventHandler(this.BtnAlterarArquivo_Click);
            // 
            // excluirToolStripMenuItem
            // 
            this.excluirToolStripMenuItem.Name = "excluirToolStripMenuItem";
            this.excluirToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.excluirToolStripMenuItem.Text = "Excluir";
            this.excluirToolStripMenuItem.Click += new System.EventHandler(this.BtnExcluir_Click);
            // 
            // atualizarToolStripMenuItem
            // 
            this.atualizarToolStripMenuItem.Name = "atualizarToolStripMenuItem";
            this.atualizarToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.atualizarToolStripMenuItem.Text = "Atualizar";
            this.atualizarToolStripMenuItem.Click += new System.EventHandler(this.BtnAtualizar_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.Controls.Add(this.listViewModelos, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 24);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(628, 328);
            this.tableLayoutPanel1.TabIndex = 5;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnNovo,
            this.btnEditar,
            this.btnBaixarModelo,
            this.btnAtualizarModelo,
            this.btnExcluir,
            this.btnAtualizar});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(628, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuPrincipal";
            // 
            // btnNovo
            // 
            this.btnNovo.Name = "btnNovo";
            this.btnNovo.Size = new System.Drawing.Size(48, 20);
            this.btnNovo.Text = "Novo";
            this.btnNovo.Click += new System.EventHandler(this.BtnNovo_Click);
            // 
            // btnEditar
            // 
            this.btnEditar.Name = "btnEditar";
            this.btnEditar.Size = new System.Drawing.Size(49, 20);
            this.btnEditar.Text = "Editar";
            this.btnEditar.Click += new System.EventHandler(this.BtnEditar_Click);
            // 
            // btnBaixarModelo
            // 
            this.btnBaixarModelo.Name = "btnBaixarModelo";
            this.btnBaixarModelo.Size = new System.Drawing.Size(94, 20);
            this.btnBaixarModelo.Text = "Baixar Modelo";
            this.btnBaixarModelo.Click += new System.EventHandler(this.BtnBaixarArquivo_Click);
            // 
            // btnAtualizarModelo
            // 
            this.btnAtualizarModelo.Name = "btnAtualizarModelo";
            this.btnAtualizarModelo.Size = new System.Drawing.Size(109, 20);
            this.btnAtualizarModelo.Text = "Atualizar Modelo";
            this.btnAtualizarModelo.Click += new System.EventHandler(this.BtnAlterarArquivo_Click);
            // 
            // btnExcluir
            // 
            this.btnExcluir.Name = "btnExcluir";
            this.btnExcluir.Size = new System.Drawing.Size(53, 20);
            this.btnExcluir.Text = "Excluir";
            this.btnExcluir.Click += new System.EventHandler(this.BtnExcluir_Click);
            // 
            // btnAtualizar
            // 
            this.btnAtualizar.Name = "btnAtualizar";
            this.btnAtualizar.Size = new System.Drawing.Size(65, 20);
            this.btnAtualizar.Text = "Atualizar";
            this.btnAtualizar.Click += new System.EventHandler(this.BtnAtualizar_Click);
            // 
            // FrmModelos
            // 
            this.ClientSize = new System.Drawing.Size(628, 352);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(563, 350);
            this.Name = "FrmModelos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Modelos de Certificado";
            this.contextMenuStrip1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem btnNovo;
        private System.Windows.Forms.ToolStripMenuItem btnEditar;
        private System.Windows.Forms.ToolStripMenuItem btnExcluir;
        private System.Windows.Forms.ToolStripMenuItem btnAtualizar;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem novoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excluirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem atualizarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem btnBaixarModelo;
        private System.Windows.Forms.ToolStripMenuItem btnAtualizarModelo;
        private System.Windows.Forms.ToolStripMenuItem baixarModeloToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem atualizarModeloToolStripMenuItem;
        private System.Windows.Forms.ColumnHeader columnHeaderColunas;
    }
}