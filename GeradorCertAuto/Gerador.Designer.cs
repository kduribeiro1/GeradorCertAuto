namespace GeradorCertAuto
{
    partial class Gerador : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Gerador()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpGeradorCertAuto = this.Factory.CreateRibbonGroup();
            this.btnEditarColunas = this.Factory.CreateRibbonButton();
            this.btnGerarExcel = this.Factory.CreateRibbonButton();
            this.btnParametros = this.Factory.CreateRibbonButton();
            this.btnCriarCert = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpGeradorCertAuto.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpGeradorCertAuto);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpGeradorCertAuto
            // 
            this.grpGeradorCertAuto.Items.Add(this.btnEditarColunas);
            this.grpGeradorCertAuto.Items.Add(this.btnGerarExcel);
            this.grpGeradorCertAuto.Items.Add(this.btnParametros);
            this.grpGeradorCertAuto.Items.Add(this.btnCriarCert);
            this.grpGeradorCertAuto.Label = "Gerador Cert Auto";
            this.grpGeradorCertAuto.Name = "grpGeradorCertAuto";
            // 
            // btnEditarColunas
            // 
            this.btnEditarColunas.Label = "Editar Colunas";
            this.btnEditarColunas.Name = "btnEditarColunas";
            this.btnEditarColunas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnEditarColunas_Click);
            // 
            // btnGerarExcel
            // 
            this.btnGerarExcel.Label = "Gerar Tabela";
            this.btnGerarExcel.Name = "btnGerarExcel";
            this.btnGerarExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGerarExcel_Click);
            // 
            // btnParametros
            // 
            this.btnParametros.Label = "Mostrar Parâmetros";
            this.btnParametros.Name = "btnParametros";
            this.btnParametros.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnParametros_Click);
            // 
            // btnCriarCert
            // 
            this.btnCriarCert.Label = "Gerar Certificados";
            this.btnCriarCert.Name = "btnCriarCert";
            this.btnCriarCert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCriarCert_Click);
            // 
            // Gerador
            // 
            this.Name = "Gerador";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Gerador_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpGeradorCertAuto.ResumeLayout(false);
            this.grpGeradorCertAuto.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpGeradorCertAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGerarExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCriarCert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnParametros;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditarColunas;
    }

    partial class ThisRibbonCollection
    {
        internal Gerador Gerador
        {
            get { return this.GetRibbon<Gerador>(); }
        }
    }
}
