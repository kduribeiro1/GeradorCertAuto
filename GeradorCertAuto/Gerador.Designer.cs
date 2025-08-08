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
            this.tabGerador = this.Factory.CreateRibbonTab();
            this.grpSisModelos = this.Factory.CreateRibbonGroup();
            this.btnAbrirSistema = this.Factory.CreateRibbonButton();
            this.btnModelosCads = this.Factory.CreateRibbonButton();
            this.grpCertSelect = this.Factory.CreateRibbonGroup();
            this.dpdModelos = this.Factory.CreateRibbonDropDown();
            this.btnCertSelectEdtCol = this.Factory.CreateRibbonButton();
            this.btnCertSelectGerTab = this.Factory.CreateRibbonButton();
            this.btnCertSelectShowPar = this.Factory.CreateRibbonButton();
            this.btnCertSelectGerCert = this.Factory.CreateRibbonButton();
            this.btnCertSelectSalvar = this.Factory.CreateRibbonButton();
            this.grpCertAberto = this.Factory.CreateRibbonGroup();
            this.btnEditarColunas = this.Factory.CreateRibbonButton();
            this.btnGerarExcel = this.Factory.CreateRibbonButton();
            this.btnParametros = this.Factory.CreateRibbonButton();
            this.btnCriarCert = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tabGerador.SuspendLayout();
            this.grpSisModelos.SuspendLayout();
            this.grpCertSelect.SuspendLayout();
            this.grpCertAberto.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tabGerador
            // 
            this.tabGerador.Groups.Add(this.grpSisModelos);
            this.tabGerador.Groups.Add(this.grpCertSelect);
            this.tabGerador.Groups.Add(this.grpCertAberto);
            this.tabGerador.Label = "Gerador Certificado";
            this.tabGerador.Name = "tabGerador";
            // 
            // grpSisModelos
            // 
            this.grpSisModelos.Items.Add(this.btnAbrirSistema);
            this.grpSisModelos.Items.Add(this.btnModelosCads);
            this.grpSisModelos.Label = "Sistema de Modelos";
            this.grpSisModelos.Name = "grpSisModelos";
            // 
            // btnAbrirSistema
            // 
            this.btnAbrirSistema.Label = "Abrir Gerador";
            this.btnAbrirSistema.Name = "btnAbrirSistema";
            this.btnAbrirSistema.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAbrirSistema_Click);
            // 
            // btnModelosCads
            // 
            this.btnModelosCads.Label = "Modelos Cadastrados";
            this.btnModelosCads.Name = "btnModelosCads";
            this.btnModelosCads.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnModelosCads_Click);
            // 
            // grpCertSelect
            // 
            this.grpCertSelect.Items.Add(this.dpdModelos);
            this.grpCertSelect.Items.Add(this.btnCertSelectEdtCol);
            this.grpCertSelect.Items.Add(this.btnCertSelectGerTab);
            this.grpCertSelect.Items.Add(this.btnCertSelectShowPar);
            this.grpCertSelect.Items.Add(this.btnCertSelectGerCert);
            this.grpCertSelect.Items.Add(this.btnCertSelectSalvar);
            this.grpCertSelect.Label = "Certificado Selecionado";
            this.grpCertSelect.Name = "grpCertSelect";
            // 
            // dpdModelos
            // 
            this.dpdModelos.Label = "Modelo: ";
            this.dpdModelos.Name = "dpdModelos";
            this.dpdModelos.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DpdModelos_SelectionChanged);
            // 
            // btnCertSelectEdtCol
            // 
            this.btnCertSelectEdtCol.Label = "Editar Colunas";
            this.btnCertSelectEdtCol.Name = "btnCertSelectEdtCol";
            this.btnCertSelectEdtCol.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCertSelectEdtCol_Click);
            // 
            // btnCertSelectGerTab
            // 
            this.btnCertSelectGerTab.Label = "Gerar Tabela";
            this.btnCertSelectGerTab.Name = "btnCertSelectGerTab";
            this.btnCertSelectGerTab.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCertSelectGerTab_Click);
            // 
            // btnCertSelectShowPar
            // 
            this.btnCertSelectShowPar.Label = "Mostrar Parâmetros";
            this.btnCertSelectShowPar.Name = "btnCertSelectShowPar";
            this.btnCertSelectShowPar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCertSelectShowPar_Click);
            // 
            // btnCertSelectGerCert
            // 
            this.btnCertSelectGerCert.Label = "Gerar Certificados";
            this.btnCertSelectGerCert.Name = "btnCertSelectGerCert";
            this.btnCertSelectGerCert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCertSelectGerCert_Click);
            // 
            // btnCertSelectSalvar
            // 
            this.btnCertSelectSalvar.Label = "Salvar Modelo";
            this.btnCertSelectSalvar.Name = "btnCertSelectSalvar";
            this.btnCertSelectSalvar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCertSelectSalvar_Click);
            // 
            // grpCertAberto
            // 
            this.grpCertAberto.Items.Add(this.btnEditarColunas);
            this.grpCertAberto.Items.Add(this.btnGerarExcel);
            this.grpCertAberto.Items.Add(this.btnParametros);
            this.grpCertAberto.Items.Add(this.btnCriarCert);
            this.grpCertAberto.Label = "Certificado Aberto";
            this.grpCertAberto.Name = "grpCertAberto";
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
            this.Tabs.Add(this.tabGerador);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Gerador_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabGerador.ResumeLayout(false);
            this.tabGerador.PerformLayout();
            this.grpSisModelos.ResumeLayout(false);
            this.grpSisModelos.PerformLayout();
            this.grpCertSelect.ResumeLayout(false);
            this.grpCertSelect.PerformLayout();
            this.grpCertAberto.ResumeLayout(false);
            this.grpCertAberto.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabGerador;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSisModelos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbrirSistema;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCertAberto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditarColunas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGerarExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnParametros;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCriarCert;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCertSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCertSelectGerTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCertSelectShowPar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCertSelectGerCert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelosCads;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCertSelectEdtCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCertSelectSalvar;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dpdModelos;
    }

    partial class ThisRibbonCollection
    {
        internal Gerador Gerador
        {
            get { return this.GetRibbon<Gerador>(); }
        }
    }
}
