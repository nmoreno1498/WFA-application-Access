using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Camada_Model;

namespace Biblioteca_DAO
{
    public class Preferencias_DAO : DAO_DAO
    {
        Preferencias_VO objPreferenciasVO;
        OleDbCommand objComando;
        OleDbDataAdapter objAdaptador;
        OleDbDataReader objLeitorDB;
        DataTable objTabela;

        public bool gerarAccess(string strNomeCompletoDaPlanilha)
        {
            try
            {
                bool resultado = false;
                abreConexao();

                StringBuilder strSQL = new StringBuilder();

                strSQL.Append(" SELECT");
                strSQL.Append(" ID");
                strSQL.Append(" ,Descricao");
                strSQL.Append(" INTO");
                strSQL.Append(" [Excel 8.0;DATABASE=" + strNomeCompletoDaPlanilha + "].[Exportar Excel]");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                objComando = new OleDbCommand(strSQL.ToString(),getConexao());

                if (objComando.ExecuteNonQuery()> 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }

                return resultado;

            }
            catch (Exception ex)
            {
                
                throw new Exception("Falhas ao gerar o Access " + ex.Message);
            }
        }

        public List<string> ImportaBD()
        {
            try
            {
                List<string> resultado = new List<string>();

                abreConexao();

                StringBuilder strSQL = new StringBuilder();

                strSQL.Append(" SELECT");
                strSQL.Append(" ID");
                strSQL.Append(" ,Descricao");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                objComando = new OleDbCommand(strSQL.ToString(),getConexao());

                objLeitorDB = objComando.ExecuteReader();

                while (objLeitorDB.Read())
                {
                    resultado.Add(objLeitorDB["Descricao"].ToString());
                }
                objLeitorDB.Close();

                return resultado;
            }
            catch (Exception ex)
            {
                
                throw new Exception("Falha ao Importar Banco De Dados " + ex.Message);
            }
        }
        public List<string> ImportaBDDesconectado()
        {
            try
            {
                List<string> resultado = new List<string>();

                abreConexao();

                StringBuilder strSQL = new StringBuilder();

                strSQL.Append(" SELECT");
                strSQL.Append(" ID");
                strSQL.Append(" ,Descricao");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                objAdaptador = new OleDbDataAdapter();
                objAdaptador.SelectCommand = objComando;

                objTabela = new DataTable();

                objAdaptador.Fill(objTabela);

                foreach (DataRow objLinha in objTabela.Rows)
                {
                    resultado.Add(objLinha["Descricao"].ToString());
                }

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha ao Importar Banco De Dados " + ex.Message);
            }
        }


        public override DataTable ConsultarBD(Object objparPreferenciasVO)
        {
            try
            {
                objPreferenciasVO = (Preferencias_VO)objparPreferenciasVO;

                StringBuilder strSQL = new StringBuilder();

                if (objPreferenciasVO.getiD() > 0)
                {
                    strSQL.Append(" SELECT");
                    strSQL.Append(" ID");
                    strSQL.Append(" ,Descricao");
                    strSQL.Append(" FROM");
                    strSQL.Append(" Preferencias_3");
                    strSQL.Append(" WHERE ID = :parID");

                    objComando = new OleDbCommand(strSQL.ToString(),getConexao());
                    objComando.Parameters.AddWithValue("parID", objPreferenciasVO.ID);
                    
                }
                else if (string.IsNullOrEmpty(objPreferenciasVO.Descricao))
                {
                    strSQL.Append(" SELECT");
                    strSQL.Append(" ID");
                    strSQL.Append(" ,Descricao");
                    strSQL.Append(" FROM");
                    strSQL.Append(" Preferencias_3");

                    objComando = new OleDbCommand(strSQL.ToString(), getConexao());
                }
                else
                {
                    strSQL.Append(" SELECT");
                    strSQL.Append(" ID");
                    strSQL.Append(" ,Descricao");
                    strSQL.Append(" FROM");
                    strSQL.Append(" Preferencias_3");
                    strSQL.Append(" WHERE Descricao = :parDescricao");

                    objComando = new OleDbCommand(strSQL.ToString(), getConexao());
                    objComando.Parameters.AddWithValue("parDescricao", objPreferenciasVO.Descricao);    
                }

                objAdaptador = new OleDbDataAdapter();
                objAdaptador.SelectCommand = objComando;

                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;
            }
            catch (Exception ex)
            {
                
                throw new Exception("Falha ao Consultar Banco De Dados De Preferencias " + ex.Message);
            }
        }
        public override void ConsultarBD(ref Object objparPreferenciasVO)
        {
            try
            {
                objPreferenciasVO = (Preferencias_VO)objparPreferenciasVO;

                StringBuilder strSQL = new StringBuilder();

                if (objPreferenciasVO.getiD() > 0)
                {
                    strSQL.Append(" SELECT");
                    strSQL.Append(" ID");
                    strSQL.Append(" ,Descricao");
                    strSQL.Append(" FROM");
                    strSQL.Append(" Preferencias_3");
                    strSQL.Append(" WHERE ID = :parID");

                    objComando = new OleDbCommand(strSQL.ToString(), getConexao());
                    objComando.Parameters.AddWithValue("parID", objPreferenciasVO.ID);

                }
                else if (string.IsNullOrEmpty(objPreferenciasVO.Descricao))
                {
                    strSQL.Append(" SELECT");
                    strSQL.Append(" ID");
                    strSQL.Append(" ,Descricao");
                    strSQL.Append(" FROM");
                    strSQL.Append(" Preferencias_3");

                    objComando = new OleDbCommand(strSQL.ToString(), getConexao());
                }
                else
                {
                    strSQL.Append(" SELECT");
                    strSQL.Append(" ID");
                    strSQL.Append(" ,Descricao");
                    strSQL.Append(" FROM");
                    strSQL.Append(" Preferencias_3");
                    strSQL.Append(" WHERE Descricao = :parDescricao");

                    objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                    objComando.Parameters.AddWithValue("parDescricao", objPreferenciasVO.Descricao);
                }

                objAdaptador = new OleDbDataAdapter();
                objAdaptador.SelectCommand = objComando;

                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                foreach (DataRow objItemPreferencias in objTabela.Rows)
                {
                    Preferencias_VO objitemPreferenciasVO = new Preferencias_VO(Convert.ToInt32(objItemPreferencias["ID"].ToString()),objItemPreferencias["Descricao"].ToString());
                    objPreferenciasVO.objPreferenciasVOCollection.Add(objitemPreferenciasVO);
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Falha ao Consultar Banco De Dados Desconectado Da Preferencias " + ex.Message);
            }
        }
        public override bool InserirBD(Object objparPreferenciasVO)
        {
            try
            {
                abreConexao();

                objPreferenciasVO = (Preferencias_VO)objparPreferenciasVO;

                bool resultado = false;

                StringBuilder strSQL = new StringBuilder();
                
                strSQL.Append(" INSERT INTO");
                strSQL.Append(" Preferencias_3(");
                strSQL.Append(" Descricao)");
                strSQL.Append(" VALUES(");
                strSQL.Append(" :parDescricao)");

                objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                objComando.Parameters.AddWithValue("parDescricao", objPreferenciasVO.Descricao);

                if (objComando.ExecuteNonQuery() > 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha ao Inserir no Banco De Dados De Preferencias " + ex.Message);
            }
        }
        public override bool ExcluirBD(Object objparPreferenciasVO)
        {
            try
            {
                abreConexao();

                objPreferenciasVO = (Preferencias_VO)objparPreferenciasVO;

                bool resultado = false;

                StringBuilder strSQL = new StringBuilder();

                strSQL.Append(" DELETE FROM");
                strSQL.Append(" Preferencias_3");
                strSQL.Append(" WHERE ID = :parID");

                objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                objComando.Parameters.AddWithValue("parID", objPreferenciasVO.ID);

                if (objComando.ExecuteNonQuery() > 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha ao Excluir no Banco De Dados De Preferencias " + ex.Message);
            }
        }
        public override bool AlterarBD(Object objparPreferenciasVO)
        {
            try
            {
                abreConexao();

                objPreferenciasVO = (Preferencias_VO)objparPreferenciasVO;

                bool resultado = false;

                StringBuilder strSQL = new StringBuilder();

                strSQL.Append(" UPDATE");
                strSQL.Append(" Preferencias_3");
                strSQL.Append(" SET");
                strSQL.Append(" Descricao = :parDescricao");
                strSQL.Append(" WHERE ID = :parID");

                objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                objComando.Parameters.AddWithValue("parDescricao", objPreferenciasVO.Descricao);
                objComando.Parameters.AddWithValue("parID", objPreferenciasVO.ID);

                if (objComando.ExecuteNonQuery() > 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha ao Alterar no Banco De Dados De Preferencias " + ex.Message);
            }
        }
    }
}
