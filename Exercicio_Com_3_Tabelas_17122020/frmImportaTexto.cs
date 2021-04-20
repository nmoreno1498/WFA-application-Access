using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Camada_Bussiness_BLL;
using Camada_Model;
using Email = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;



namespace Exercicio_Com_3_Tabelas_17122020
{
    public partial class frmImportaTexto : Form
    {
        Preferencias_BLL objPreferenciasBLL;
        Preferencias_VO objPreferenciasVO;
        Preferencias_De_Familiares_VO objPreferenciasDeFamiliaresVO;
        Preferencias_De_Familiares_BLL objPreferenciasDeFamiliaresBLL;
        OleDbCommand objComando;
        OleDbDataAdapter objAdaptador;
        OleDbDataReader objLeitor;
        DataTable objTabela;
        Familiares_BLL objFamiliaresBLL;
        Familiares_VO objFamiliaresVO;
        int intiD_Antigo, intCodFamiliar, intCodPrefFam, intIDPrefFam;
        string strValorAntigo, strValorFamiliarAntigo, strValorAntigoNome, strValorAntigoDescricao;
        bool boolInserirValor, boolInserirValorFamiliar, boolInserirValorPrefFamiliar;


        Email.Application objEmailApp;
        Email.MailItem objEmailMensagem;
        Email.OlAttachmentType objEmaiArquivosAnexos;
        string[] objEmailMatriz = new String[0];
        long objEmailPocisionInicial;
        string strEmailDisplay;

        Excel.Application objExcelApp;
        Excel.Workbook objExcelPastaDeTrabalho;
        Excel.Worksheet objExcelPlanilha;
        Excel.Range objExcelCelulas;

        public frmImportaTexto()
        {
            InitializeComponent();
        }

        private void btnDesvCond_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Escolha Aceitar ou Cancelar ","Aviso",MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show("Escolheu Aceitar ");
            }
            else
            {
                MessageBox.Show("Escolheu Cancelar ");
            }
        }

        private void btnImpTxt_Click(object sender, EventArgs e)
        {
            lstbxPreferencias.Items.Clear();

            objPreferenciasBLL = new Preferencias_BLL();

            lstbxPreferencias.Items.AddRange(objPreferenciasBLL.ImportaTextoWhile().ToArray());
        }

        private void btnImpBd_Click(object sender, EventArgs e)
        {
            lstbxPreferencias.Items.Clear();

            objPreferenciasBLL = new Preferencias_BLL();

            lstbxPreferencias.Items.AddRange(objPreferenciasBLL.ImportaBD().ToArray());
        }

        private void btnImpBdDesc_Click(object sender, EventArgs e)
        {
            lstbxPreferencias.Items.Clear();

            objPreferenciasBLL = new Preferencias_BLL();

            lstbxPreferencias.Items.AddRange(objPreferenciasBLL.ImportaBDDseconectado().ToArray());
        }

        private void btnConsBD_Click(object sender, EventArgs e)
        {
            ConsultarBD();
        }
        public void ConsultarBD(int ? intiD = null, string strPreferencias = null)
        {
            try
            {
                objPreferenciasVO = new Preferencias_VO();

                if (!string.IsNullOrEmpty(intiD.ToString()))
                {
                    objPreferenciasVO.ID = Convert.ToInt32(intiD);
                }
                if (!string.IsNullOrEmpty(strPreferencias))
                {
                    objPreferenciasVO.Descricao = strPreferencias;
                }

                objPreferenciasBLL = new Preferencias_BLL();

                bndsrcPreferencias.DataSource = objPreferenciasBLL.ConsultarBD(objPreferenciasVO);

                dtgvwPreferencias.DataSource = bndsrcPreferencias;
            }
            catch (Exception ex)
            {
                
                MessageBox.Show("Falhas ao Consultar no Banco De Dados ==> " + ex.Message);
            }
        }

        private void frmImportaTexto_Load(object sender, EventArgs e)
        {
            ConsultarBD();
            ConsultarBDFamiliares();
            dtgvwPrefFamRefresh();
        }

        private void btnInsBd_Click(object sender, EventArgs e)
        {
            InserirBD(dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
            ConsultarBD();
        }
        public void InserirBD(string strPreferencias)
        {
            try
            {
                objPreferenciasVO = new Preferencias_VO();
                objPreferenciasVO.Descricao = strPreferencias;

                objPreferenciasBLL = new Preferencias_BLL();

                if (objPreferenciasBLL.InserirBD(objPreferenciasVO))
                {
                    MessageBox.Show("Inclusao Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Inclusao ");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Falhas ao Inserir no Banco De Dados ==> " + ex.Message);
            }
        }
        private void btnExcBd_Click(object sender, EventArgs e)
        {
            ExcluirBD(Convert.ToInt32(dtgvwPreferencias.CurrentRow.Cells["ID"].Value.ToString()));
        }
        public void ExcluirBD(int intPreferencia)
        {
             try
            {
                objPreferenciasVO = new Preferencias_VO();
                objPreferenciasVO.ID = intPreferencia;

                objPreferenciasBLL = new Preferencias_BLL();

                if (objPreferenciasBLL.ExcluirBD(objPreferenciasVO))
                {
                    MessageBox.Show("Exclusao Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Exclusao ");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Falhas ao Excluir no Banco De Dados ==> " + ex.Message);
            }
        }

        private void btnAltBd_Click(object sender, EventArgs e)
        {
            AlterarBD(intiD_Antigo, dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
        }
        public void AlterarBD(int intiD_Preferencia, string strValorPreferencia)
        {
            try
            {
                objPreferenciasVO = new Preferencias_VO(intiD_Preferencia, strValorPreferencia);

                objPreferenciasBLL = new Preferencias_BLL();

                if (objPreferenciasBLL.AlterarBD(objPreferenciasVO))
                {
                    MessageBox.Show("Alteraçao Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Alteraçao ");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Falhas ao Alterar no Banco De Dados ==> " + ex.Message);
            }
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            boolInserirValor = true;
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir " + strValorAntigo,"Aviso",MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                ExcluirBD(intiD_Antigo);
            }
            ConsultarBD();
        }

        private void bndnavbtnConfPref_Click(object sender, EventArgs e)
        {
            if (boolInserirValor)
            {
                if (MessageBox.Show("Deseja Inserir " + dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),"Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    InserirBD(dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
                }
                boolInserirValor = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Excluir " + strValorAntigo + "Para a Preferencia" + dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    AlterarBD(intiD_Antigo,dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
                }
            }
            ConsultarBD();
        }

        private void bndnavbtnConsPref_Click(object sender, EventArgs e)
        {
            ConsultarBD(null, bndnavtxtPref.Text);
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            gerarExcel(dtgvwPreferencias);
        }
        public void gerarExcel(DataGridView objparDtgvw)
        {
            //algoritmo para criaçao e exportaçao de Excel
            try
            {
                objExcelApp = new Excel.Application();
                objExcelApp.Visible = true;

                objExcelPastaDeTrabalho = objExcelApp.Workbooks.Add();

                objExcelPlanilha = objExcelPastaDeTrabalho.Worksheets[1];

                int coluna = 1, linha = 2, linhaDoCabeçalho = 1;

                objExcelCelulas = objExcelPlanilha.Cells[linha, coluna];

                Excel.Range objExcelCabeçalho = objExcelPlanilha.Cells[linhaDoCabeçalho, coluna];

                foreach (DataGridViewRow objLinhasTabela in objparDtgvw.Rows)
                {
                    foreach (DataGridViewColumn objColunasTabela in objparDtgvw.Columns)
                    {
                        if (linha <= 2)
                        {
                            objExcelCabeçalho.set_Value(Type.Missing, objColunasTabela.HeaderText.ToString());
                        }
                        if (objLinhasTabela.Cells[coluna - 1].Value != null)
                        {
                            objExcelCelulas.set_Value(Type.Missing, objLinhasTabela.Cells[coluna - 1].Value.ToString());
                        }
                        coluna++;
                        if (linha <= 2)
                        {
                            objExcelCabeçalho = objExcelPlanilha.Cells[linhaDoCabeçalho, coluna];
                        }
                        objExcelCelulas = objExcelPlanilha.Cells[linha, coluna];
                    }
                    objExcelCelulas = objExcelPlanilha.Cells[linha, coluna];
                    linha++;
                    coluna = 1;
                }
                objExcelPastaDeTrabalho.SaveAs(@"D:\Curso Programacao\Importaciones Excel\TreinoExcelDePreferencias1.xlsx",
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlShared);
                objExcelApp.Quit();
                MessageBox.Show("Geraçao de Excel Concluida ", "aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {

                throw new Exception("Falhas ao gerar o Excel " + ex.Message);
            }
        }

        public void ConsultarBDFamiliares(int ? intCod = null, string strNome = null)
        {
            try
            {
                objFamiliaresBLL = new Familiares_BLL();

                objFamiliaresVO = new Familiares_VO();
                objFamiliaresVO.Cod = Convert.ToInt32(intCod == null ? 0 : intCod);
                objFamiliaresVO.Nome = strNome;

                bndsrcFamiliar.DataSource = objFamiliaresBLL.ConsultarBD(objFamiliaresVO);

                dtgvwFamiliar.Columns.Clear();
                dtgvwFamiliar.DataSource = null;
                dtgvwFamiliar.AllowUserToAddRows = false;

                dtgvwFamiliar.Columns.Add("Cod","Codigo Do Familiar");
                dtgvwFamiliar.Columns["Cod"].DataPropertyName = "Cod";

                dtgvwFamiliar.Columns.Add("Nome", "Nome Do Familiar");
                dtgvwFamiliar.Columns["Nome"].DataPropertyName = "Nome";

                DataGridViewComboBoxColumn objColunaComboBoxSelecionavel = new DataGridViewComboBoxColumn();
                objColunaComboBoxSelecionavel.Name = "Sexo";
                objColunaComboBoxSelecionavel.ValueType = typeof(string);
                objColunaComboBoxSelecionavel.HeaderText = "Sexo Do Familiar";
                objColunaComboBoxSelecionavel.Items.Add("MASCULINO");
                objColunaComboBoxSelecionavel.Items.Add("FEMENINO");
                objColunaComboBoxSelecionavel.Items.Add("INDEFINIDO");
                objColunaComboBoxSelecionavel.DataPropertyName = "Sexo";

                dtgvwFamiliar.Columns.Add(objColunaComboBoxSelecionavel);

                dtgvwFamiliar.Columns.Add("Idade", "Idade Do Familiar");
                dtgvwFamiliar.Columns["Idade"].DataPropertyName = "Idade";

                dtgvwFamiliar.Columns.Add("GanhoTotalMensal", "GanhoTotalMensal Do Familiar");
                dtgvwFamiliar.Columns["GanhoTotalMensal"].DataPropertyName = "GanhoTotalMensal";

                dtgvwFamiliar.Columns.Add("GastoTotalMensal", "GastoTotalMensal Do Familiar");
                dtgvwFamiliar.Columns["GastoTotalMensal"].DataPropertyName = "GastoTotalMensal";

                dtgvwFamiliar.Columns.Add("Observaçao", "Observaçao Do Familiar");
                dtgvwFamiliar.Columns["Observaçao"].DataPropertyName = "Observaçao";

                dtgvwFamiliar.DataSource = bndsrcFamiliar;

                cmbbxFamiliar.DataSource = bndsrcFamiliar.DataSource;
                cmbbxFamiliar.DisplayMember = "Nome";
                cmbbxFamiliar.ValueMember = "Cod";
                cmbbxFamiliar.SelectedIndex = Convert.ToInt16(intCod > 0 ? intCod -1 : 0);
            }
            catch (Exception ex)
            {

                throw new Exception("Falhas ao Consultar no Banco De Dados De Familiares ==> " + ex.Message);
            }
        }
        public void InserirBDFamiliares(string strNome, string strSexo, int ? intIdade = null, double ? dbGanhoTotalMensal = null, double ? dbGastoTotalMensal = null, string strObservaçao = null)
        {
            try
            {
                objFamiliaresBLL = new Familiares_BLL();

                objFamiliaresVO = new Familiares_VO();
                objFamiliaresVO.Nome = strNome;
                objFamiliaresVO.Sexo = strSexo;
                objFamiliaresVO.Idade = Convert.ToInt32(intIdade == null ? 0 : intIdade);
                objFamiliaresVO.GanhoTotalMensal = Convert.ToDouble(dbGanhoTotalMensal == null ? 0 : dbGanhoTotalMensal);
                objFamiliaresVO.GastoTotalMensal = Convert.ToDouble(dbGastoTotalMensal == null ? 0 : dbGastoTotalMensal);
                objFamiliaresVO.Observaçao = strObservaçao;

                if (objFamiliaresBLL.InserirBD(objFamiliaresVO))
                {
                    MessageBox.Show("Inserçao De Familiar Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Inserçao De Familiar ");
                }
                
            }
            catch (Exception ex)
            {

                throw new Exception("Falhas ao Inserir no Banco De Dados De Familiares ==> " + ex.Message);
            }
        }
        public void ExcluirBDFamiliares(int intCodExcluirFamiliar)
        {
            try
            {
                objFamiliaresBLL = new Familiares_BLL();

                objFamiliaresVO = new Familiares_VO();
                objFamiliaresVO.Cod = intCodExcluirFamiliar;

                if (objFamiliaresBLL.ExcluirBD(objFamiliaresVO))
                {
                    MessageBox.Show("Exclusao De Familiar Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Exclusao De Familiar ");
                }

            }
            catch (Exception ex)
            {

                throw new Exception("Falhas ao Excluir no Banco De Dados De Familiares ==> " + ex.Message);
            }
        }
        public void AlterarBDFamiliares(int intCod, string strNome, string strSexo, int? intIdade = null, double? dbGanhoTotalMensal = null, double? dbGastoTotalMensal = null, string strObservaçao = null)
        {
            try
            {
                objFamiliaresBLL = new Familiares_BLL();

                objFamiliaresVO = new Familiares_VO();
                objFamiliaresVO.Cod = intCod;
                objFamiliaresVO.Nome = strNome;
                objFamiliaresVO.Sexo = strSexo;
                objFamiliaresVO.Idade = Convert.ToInt32(intIdade == null ? 0 : intIdade);
                objFamiliaresVO.GanhoTotalMensal = Convert.ToDouble(dbGanhoTotalMensal == null ? 0 : dbGanhoTotalMensal);
                objFamiliaresVO.GastoTotalMensal = Convert.ToDouble(dbGastoTotalMensal == null ? 0 : dbGastoTotalMensal);
                objFamiliaresVO.Observaçao = strObservaçao;

                if (objFamiliaresBLL.AlterarBD(objFamiliaresVO))
                {
                    MessageBox.Show("Alteraçao De Familiar Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Alteraçao De Familiar ");
                }

            }
            catch (Exception ex)
            {

                throw new Exception("Falhas ao Alterar no Banco De Dados De Familiares ==> " + ex.Message);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            boolInserirValorFamiliar = true;
            dtgvwFamiliar.CurrentRow.Cells["Cod"].ReadOnly = true;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir o Familiar " + strValorFamiliarAntigo,"Aviso",MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                ExcluirBDFamiliares(intCodFamiliar);
            }
            ConsultarBDFamiliares();
        }

        private void bndnavbtnConf_Click(object sender, EventArgs e)
        {
            if (boolInserirValorFamiliar)
            {
                if (MessageBox.Show("Deseja Inserir o Familiar " + dtgvwFamiliar.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    InserirBDFamiliares(dtgvwFamiliar.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                        dtgvwFamiliar.CurrentRow.Cells["Sexo"].EditedFormattedValue.ToString(),
                                        Convert.ToInt32(dtgvwFamiliar.CurrentRow.Cells["Idade"].EditedFormattedValue.ToString()),
                                        Convert.ToDouble(dtgvwFamiliar.CurrentRow.Cells["GanhoTotalMensal"].EditedFormattedValue.ToString()),
                                        Convert.ToDouble(dtgvwFamiliar.CurrentRow.Cells["GastoTotalMensal"].EditedFormattedValue.ToString()),
                                        dtgvwFamiliar.CurrentRow.Cells["Observaçao"].EditedFormattedValue.ToString());
                }
                boolInserirValorFamiliar = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar o Familiar " + strValorFamiliarAntigo + " Para o Novo Valor " + dtgvwFamiliar.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    AlterarBDFamiliares(Convert.ToInt32(dtgvwFamiliar.CurrentRow.Cells["Cod"].EditedFormattedValue.ToString()),
                                        dtgvwFamiliar.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                        dtgvwFamiliar.CurrentRow.Cells["Sexo"].EditedFormattedValue.ToString(),
                                        Convert.ToInt32(dtgvwFamiliar.CurrentRow.Cells["Idade"].EditedFormattedValue.ToString()),
                                        Convert.ToDouble(dtgvwFamiliar.CurrentRow.Cells["GanhoTotalMensal"].EditedFormattedValue.ToString()),
                                        Convert.ToDouble(dtgvwFamiliar.CurrentRow.Cells["GastoTotalMensal"].EditedFormattedValue.ToString()),
                                        dtgvwFamiliar.CurrentRow.Cells["Observaçao"].EditedFormattedValue.ToString());
                }
            }
            ConsultarBDFamiliares();
        }

        private void bndnavbtnConsultaFam_Click(object sender, EventArgs e)
        {
            ConsultarBDFamiliares(null, bndnavtxtFam.Text);
        }

        private void btnGerarExcel_Click(object sender, EventArgs e)
        {
            gerarExcel(dtgvwFamiliar);
        }

        private void dtgvwFamiliar_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strValorFamiliarAntigo = dtgvwFamiliar.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString();
            if (!string.IsNullOrEmpty(dtgvwFamiliar.CurrentRow.Cells["Cod"].Value.ToString()))
            {
                intCodFamiliar = Convert.ToInt32(dtgvwFamiliar.CurrentRow.Cells["Cod"].Value.ToString());
            }
        }
        public void dtgvwPrefFamRefresh()
        {
            try
            {
                ConsultarBDPreferenciaDeFamiliar(Convert.ToInt32(cmbbxFamiliar.SelectedValue.ToString()),null,cmbbxFamiliar.Text,null);
            }
            catch (Exception ex)
            {
                
                throw new Exception("Falha ao Dar o Refresh " + ex.Message);
            }
        }
        public void ConsultarBDPreferenciaDeFamiliar(int ? intCod = null, int ? intiD = null, string strNome = null, string strDescricao = null)
        {
            try
            {
                objPreferenciasDeFamiliaresBLL = new Preferencias_De_Familiares_BLL();

                objPreferenciasDeFamiliaresVO = new Preferencias_De_Familiares_VO();
                objFamiliaresVO = new Familiares_VO();
                objFamiliaresVO.Cod = Convert.ToInt32(intCod == null ? 0 : intCod);
                objFamiliaresVO.Nome = strNome;

                objPreferenciasVO = new Preferencias_VO();
                objPreferenciasVO.ID = Convert.ToInt32(intiD == null ? 0 : intiD);
                objPreferenciasVO.Descricao = strDescricao;

                objPreferenciasDeFamiliaresVO.ObjFamiliarVO = objFamiliaresVO;
                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO = objPreferenciasVO;

                bndsrcPrefFam.DataSource = objPreferenciasDeFamiliaresBLL.ConsultarBD(objPreferenciasDeFamiliaresVO);

                dtgvwPrefFam.Columns.Clear();
                dtgvwPrefFam.DataSource = null;
                dtgvwPrefFam.AllowUserToAddRows = false;

                Preferencias_BLL objPreferencias = new Preferencias_BLL();
                bndsrcPrefFamLookUp.DataSource = objPreferenciasBLL.ConsultarBD(objPreferenciasVO);

                DataGridViewComboBoxColumn objColunaComboFamiliarLookUp = new DataGridViewComboBoxColumn();
                objColunaComboFamiliarLookUp.DataSource = bndsrcFamiliar.DataSource;
                objColunaComboFamiliarLookUp.HeaderText = "Codigo Do Familiar";
                objColunaComboFamiliarLookUp.Name = "Cod";
                objColunaComboFamiliarLookUp.DisplayMember = "Nome";
                objColunaComboFamiliarLookUp.ValueType = typeof(int);
                objColunaComboFamiliarLookUp.ValueMember = "Cod";
                objColunaComboFamiliarLookUp.DataPropertyName = "Cod";
                dtgvwPrefFam.Columns.Add(objColunaComboFamiliarLookUp);
                dtgvwPrefFam.Columns["Cod"].ValueType = typeof(int);

                DataGridViewComboBoxColumn objColunaComboPreferenciaLookUp = new DataGridViewComboBoxColumn();
                objColunaComboPreferenciaLookUp.DataSource = bndsrcPrefFamLookUp.DataSource;
                objColunaComboPreferenciaLookUp.HeaderText = "Descricao Do Familiar";
                objColunaComboPreferenciaLookUp.Name = "ID";
                objColunaComboPreferenciaLookUp.ValueType = typeof(int);
                objColunaComboPreferenciaLookUp.ValueMember = "ID";
                objColunaComboPreferenciaLookUp.DisplayMember = "Descricao";
                dtgvwPrefFam.Columns.Add(objColunaComboPreferenciaLookUp);
                objColunaComboPreferenciaLookUp.DataPropertyName = "ID";


                dtgvwPrefFam.Columns["ID"].ValueType = typeof(int);

                dtgvwPrefFam.Columns.Add("Intensidade", "Intensidade Do Familiar");
                dtgvwPrefFam.Columns["Intensidade"].DataPropertyName = "Intensidade";

                dtgvwPrefFam.Columns.Add("Observaçao", "Observaçao Do Familiar");
                dtgvwPrefFam.Columns["Observaçao"].DataPropertyName = "Observaçao";

                dtgvwPrefFam.DataSource = bndsrcPrefFam;

                cmbbxPrefFam.Items.Clear();

                foreach (DataRow objPreferenciaLinha in ((DataTable)bndsrcPreferencias.DataSource).Rows)
                {
                    cmbbxPrefFam.Items.Add(objPreferenciaLinha["ID"].ToString() + "-" + objPreferenciaLinha["Descricao"].ToString());
                }

            }
            catch (Exception ex)
            {
                
                MessageBox.Show("Falhas ao Consultar Banco De Dados De Preferencia De Familiares " + ex.Message);
            }
        }
        public void InserirBDPreferenciaDeFamiliar(int intCod, int intiD, float fltIntensidade, string strNome = null, string strDescricao = null, string strObservaçao = null)
        {
            try
            {
                objPreferenciasDeFamiliaresVO = new Preferencias_De_Familiares_VO();

                objPreferenciasDeFamiliaresVO.ObjFamiliarVO = new Familiares_VO();
                objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod = intCod;
                objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Nome = strNome;

                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO = new Preferencias_VO();
                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID = intiD;
                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.Descricao = strDescricao;

                objPreferenciasDeFamiliaresVO.Intensidade = fltIntensidade;
                objPreferenciasDeFamiliaresVO.Observaçao = strObservaçao;

                objPreferenciasDeFamiliaresBLL = new Preferencias_De_Familiares_BLL();

                if (objPreferenciasDeFamiliaresBLL.InserirBD(objPreferenciasDeFamiliaresVO))
                {
                    MessageBox.Show("Inserçao De Preferencia De Familiar Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Inserçao De Preferencia De Familiar");
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("Falhas ao Inserir Banco De Dados De Preferencia De Familiares " + ex.Message);
            }
        }
        public void ExcluirBDPreferenciaDeFamiliares(int intCodPrefFamExc, int intIDPrefFamExc)
        {
            try
            {
                objPreferenciasDeFamiliaresVO = new Preferencias_De_Familiares_VO();

                objPreferenciasDeFamiliaresVO.ObjFamiliarVO = new Familiares_VO();
                objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod = intCodPrefFamExc;

                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO = new Preferencias_VO();
                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID= intIDPrefFamExc;

                objPreferenciasDeFamiliaresBLL = new Preferencias_De_Familiares_BLL();

                if (objPreferenciasDeFamiliaresBLL.ExcluirBD(objPreferenciasDeFamiliaresVO))
                {
                    MessageBox.Show("Exclusao De Preferencia De Familiar Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Exclusao De Preferencia De Familiar");
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("Falhas ao Excluir Banco De Dados De Preferencia De Familiares " + ex.Message);
            }
        }
        public void AlterarBDPreferenciaDeFamiliar(int intCod , int intiD, float fltIntensidade, string strNome = null, string strDescricao = null, string strObservaçao = null)
        {
            try
            {
                objPreferenciasDeFamiliaresVO = new Preferencias_De_Familiares_VO();

                objPreferenciasDeFamiliaresVO.ObjFamiliarVO = new Familiares_VO();
                objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod = intCod;
                objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Nome = strNome;

                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO = new Preferencias_VO();
                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID = intiD;
                objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.Descricao = strDescricao;

                objPreferenciasDeFamiliaresVO.Intensidade = fltIntensidade;
                objPreferenciasDeFamiliaresVO.Observaçao = strObservaçao;

                objPreferenciasDeFamiliaresBLL = new Preferencias_De_Familiares_BLL();

                if (objPreferenciasDeFamiliaresBLL.AlterarBD(objPreferenciasDeFamiliaresVO))
                {
                    MessageBox.Show("Alteraçao De Preferencia De Familiar Realizada ");
                }
                else
                {
                    MessageBox.Show("Problemas na Alteraçao De Preferencia De Familiar");
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("Falhas ao Alterar Banco De Dados De Preferencia De Familiares " + ex.Message);
            }
        }

        private void dtgvwPrefFam_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgvwPrefFam.CurrentRow.Cells["Cod"].EditedFormattedValue.ToString()))
            {
                intIDPrefFam = Convert.ToInt32(dtgvwPrefFam.CurrentRow.Cells["ID"].Value.ToString());

                strValorAntigoDescricao = dtgvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString();

                intCodPrefFam = Convert.ToInt32(dtgvwPrefFam.CurrentRow.Cells["Cod"].Value.ToString());

                strValorAntigoNome = dtgvwPrefFam.CurrentRow.Cells["Cod"].EditedFormattedValue.ToString();

                dtgvwPrefFam.CurrentRow.Cells["ID"].Selected = false;

                dtgvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = true;

                dtgvwPrefFam.CurrentRow.Cells["Cod"].Selected = false;

                dtgvwPrefFam.CurrentRow.Cells["Cod"].ReadOnly = true;
            }
            else
            {
                dtgvwPrefFam.CurrentRow.Cells["ID"].Selected = true;

                dtgvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = false;

                dtgvwPrefFam.CurrentRow.Cells["Cod"].Selected = false;

                dtgvwPrefFam.CurrentRow.Cells["Cod"].ReadOnly = true;
            }
        }

        private void cmbbxFamiliar_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(((ComboBox)sender).SelectedText))
            {
                dtgvwPrefFamRefresh();
            }
        }

        private void cmbbxFamiliar_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(((ComboBox)sender).ValueMember.ToString()))
            {
                dtgvwPrefFamRefresh();
            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            boolInserirValorPrefFamiliar = true;
            dtgvwPrefFam.CurrentRow.Cells["Cod"].ReadOnly = true;
            dtgvwPrefFam.CurrentRow.Cells["Cod"].Selected = false;
            dtgvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = false;
            dtgvwPrefFam.CurrentRow.Cells["ID"].Selected = true;

        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir a Preferencia De Familiar " + strValorAntigoNome + " Com a sua Preferencia " + strValorAntigoDescricao,"Aviso",MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                ExcluirBDPreferenciaDeFamiliares(intCodPrefFam, intIDPrefFam);
            }
            dtgvwPrefFamRefresh();
        }

        private void bndnavbtnConfPrefFam_Click(object sender, EventArgs e)
        {
            if (boolInserirValorPrefFamiliar)
            {
                if (MessageBox.Show("Deseja Inserir a Preferencia De Familiar " + cmbbxFamiliar.Text + " Para a Preferencia " + dtgvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    InserirBDPreferenciaDeFamiliar(Convert.ToInt32(cmbbxFamiliar.SelectedValue.ToString()),
                                        Convert.ToInt32(dtgvwPrefFam.CurrentRow.Cells["ID"].Value.ToString()),
                                        Convert.ToSingle(dtgvwPrefFam.CurrentRow.Cells["Intensidade"].Value.ToString()),
                                        cmbbxFamiliar.Text.ToString(),
                                        dtgvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString(),
                                        dtgvwPrefFam.CurrentRow.Cells["Observaçao"].Value.ToString());
                }
                boolInserirValorPrefFamiliar = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar a Preferencia De Familiar " + strValorAntigoNome + " Com a sua Preferencia " + strValorAntigoDescricao, "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    AlterarBDPreferenciaDeFamiliar(intCodPrefFam,intIDPrefFam,
                                                   Convert.ToSingle(dtgvwPrefFam.CurrentRow.Cells["Intensidade"].EditedFormattedValue.ToString()),
                                                   cmbbxFamiliar.Text.ToString(),
                                                   dtgvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString(),
                                                   dtgvwPrefFam.CurrentRow.Cells["Observaçao"].EditedFormattedValue.ToString());
                }
                dtgvwPrefFamRefresh();
            }
        }

        private void bndnavbtnConsPrefFam_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbbxPrefFam.Text.Trim()))
            {
                ConsultarBDPreferenciaDeFamiliar(Convert.ToInt32(cmbbxFamiliar.SelectedValue.ToString()),
                    Convert.ToInt32(cmbbxPrefFam.Text.Substring(0, cmbbxPrefFam.Text.IndexOf("-"))),
                    cmbbxFamiliar.Text,
                    cmbbxPrefFam.Text.Substring(cmbbxPrefFam.Text.IndexOf("-") + 1));
            }
            else
            {
                dtgvwPrefFamRefresh();
            }
        }

        private void btngerarExcelPrefFm_Click(object sender, EventArgs e)
        {
            gerarExcel(dtgvwPrefFam);
        }

        private void dtgvwPreferencias_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strValorAntigo = dtgvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString();

            if (!string.IsNullOrEmpty(dtgvwPreferencias.CurrentRow.Cells["ID"].Value.ToString()))
            {
                intiD_Antigo = Convert.ToInt32(dtgvwPreferencias.CurrentRow.Cells["ID"].Value.ToString());
            }
        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            gerarEmail();
        }
        public void gerarEmail()
        {
            try
            {
                objEmailApp = new Email.Application();

                objEmailMensagem = objEmailApp.CreateItem(Email.OlItemType.olMailItem);

                objEmailMensagem.SentOnBehalfOfName = "nmoreno1498@gmail.com";
                objEmailMensagem.To = "robertosptcosta@gmail.com";
                objEmailMensagem.CC = "nmoreno1498@gmail.com";
                objEmailMensagem.BCC = "nmoreno1498@gmail.com";
                objEmailMensagem.Subject = "Testando automaçao excel";
                objEmailMensagem.Body = "Olà! Esse è um teste de Automaçao do Email via Outlook" +
                                                Environment.NewLine +
                                                "Conforme Combinado segue esse texto para o envio de email pelo C# \n" +
                                                "(Esse email é automático, não responda)";

                if (MessageBox.Show("Deseja anexar arquivos","aviso",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    ofdEscolheAnexos.Title = "Escolha arquivos a serem anexados";
                    ofdEscolheAnexos.InitialDirectory = @"D:\";
                    ofdEscolheAnexos.ShowDialog();

                    if(!string.IsNullOrEmpty(ofdEscolheAnexos.FileName))
                    {
                        objEmailMatriz = ofdEscolheAnexos.FileName.Split(';');

                        for(int i = 0; i < objEmailMatriz.Length; i++)
                        {
                            objEmaiArquivosAnexos = Email.OlAttachmentType.olByValue;

                            objEmailPocisionInicial = objEmailMensagem.Body.Length;

                            strEmailDisplay = objEmailMatriz[i].ToString();

                            objEmailMensagem.Attachments.Add(objEmailMatriz[i].ToString(), objEmaiArquivosAnexos, objEmailPocisionInicial, strEmailDisplay);
                            //objEmailMensagem.Attachments.Add(objEmailMatriz[i],objEmaiArquivosAnexos,objEmailPocisionInicial,strEmailDisplay);
                        }
                    }
                }
                if (MessageBox.Show("Envia o email com confirmaçao e visualizaçao ","aviso",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    objEmailMensagem.Display();
                }
                else
                {
                    objEmailMensagem.Send();
                    MessageBox.Show("E-mail enviado com confirmaçao e visualizaçao ","aviso",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                }
                MessageBox.Show("Finalizaçao do envio de E-mail ","aviso",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                
                throw new Exception("Falhas na geraçao do E-mail " + ex.Message);
            }
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            gerarDataBase(new Preferencias_BLL().ConsultarBD(new Preferencias_VO()));
            //objeto anonimo
        }
        public void gerarDataBase(DataTable objTabela)
        {
            try
            {
                objExcelApp = new Excel.Application();
                objExcelApp.Visible = true;

                objExcelPastaDeTrabalho = objExcelApp.Workbooks.Add();

                objExcelPlanilha = objExcelPastaDeTrabalho.Worksheets[1];

                int coluna = 1, linha = 2, linhaDoCabeçalho = 1;

                objExcelCelulas = objExcelPlanilha.Cells[linha, coluna];

                Excel.Range objExcelCabeçalho = objExcelPlanilha.Cells[linhaDoCabeçalho, coluna];

                foreach (DataRow objLinhasTabela in objTabela.Rows)
                {
                    foreach (DataColumn objColunasTabela in objTabela.Columns)
                    {
                        if (linha <= 2)
                        {
                            objExcelCabeçalho.set_Value(Type.Missing, objColunasTabela.ColumnName.ToString());
                        }
                        if (!string.IsNullOrEmpty(objLinhasTabela[coluna - 1].ToString()))
                        {
                            objExcelCelulas.set_Value(Type.Missing, objLinhasTabela[coluna - 1].ToString());
                        }
                        //objExcelCelulas.set_Value(Type.Missing, objLinhasTabela.Field<object>(objColunasTabela) == null ? string.Empty : objLinhasTabela.Field<Object>(objColunasTabela).ToString());
                        coluna++;
                        if (linha <= 2)
                        {
                            objExcelCabeçalho = objExcelPlanilha.Cells[linhaDoCabeçalho, coluna];
                        }
                        objExcelCelulas = objExcelPlanilha.Cells[linha, coluna];
                    }
                    objExcelCelulas = objExcelPlanilha.Cells[linha, coluna];
                    linha++;
                    coluna = 1;
                }
                objExcelPastaDeTrabalho.SaveAs(@"D:\Curso Programacao\Importaciones Excel\TreinoExcelDePreferencias1.xlsx",
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlShared);
                objExcelApp.Quit();
                MessageBox.Show("Geraçao de Excel Concluida ", "aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {

                throw new Exception("Falhas ao gerar o Excel " + ex.Message);
            }
        }


        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            gerarAccess(objPreferenciasBLL.GetType());
        }
        public void gerarAccess(Type objType)
        {
            try
            {
                sfdPlanilhaInterop.ShowDialog();

                if (objType == typeof(Preferencias_BLL))
                {
                    Preferencias_BLL objGenerico = new Preferencias_BLL();
                    objGenerico.gerarAccess(sfdPlanilhaInterop.FileName);
                }
                else if (objType == typeof(Familiares_BLL))
                {
                    Familiares_BLL objGenerico = new Familiares_BLL();
                    objGenerico.gerarAccess(sfdPlanilhaInterop.FileName);
                }
                else 
                {
                    Preferencias_De_Familiares_BLL objGenerico = new Preferencias_De_Familiares_BLL();
                    objGenerico.gerarAccess(sfdPlanilhaInterop.FileName);
                }
                MessageBox.Show("Exportaçao do Excel Realizada ", "aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {

                throw new Exception("Falhas na geraçao do Access " + ex.Message);
            }
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            gerarDataBase(new Familiares_BLL().ConsultarBD(new Familiares_VO()));
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            gerarAccess(objFamiliaresBLL.GetType());
        }

        private void toolStripButton19_Click(object sender, EventArgs e)
        {
            gerarDataBase(new Preferencias_De_Familiares_BLL().ConsultarBD(new Preferencias_De_Familiares_VO()));
        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {
            gerarAccess(objPreferenciasDeFamiliaresBLL.GetType());
        }
    }
}
