using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_Model;
using System.Data.OleDb;


namespace Biblioteca_DAO
{
   public class Familiares_DAO : DAO_DAO
    {
       OleDbCommand objComando;
       OleDbDataAdapter objAdaptador;
       DataTable objTabela;
       Familiares_VO objFamiliaresVO;

       public bool gerarAccess(string strNomeCompletoDaPlanilha)
       {
           try
           {
               bool resultado = false;
               abreConexao();

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" SELECT");
               strSQL.Append(" Cod");
               strSQL.Append(" ,Nome");
               strSQL.Append(" ,Sexo");
               strSQL.Append(" ,Idade");
               strSQL.Append(" ,GanhoTotalMensal");
               strSQL.Append(" ,GastoTotalMensal");
               strSQL.Append(" ,Observaçao");
               strSQL.Append(" INTO");
               strSQL.Append(" [Excel 8.0;DATABASE=" + strNomeCompletoDaPlanilha + "].[Exportar Excel]");
               strSQL.Append(" FROM");
               strSQL.Append(" FAMILIARES");

               objComando = new OleDbCommand(strSQL.ToString(), getConexao());

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

               throw new Exception("Falhas ao gerar o Access " + ex.Message);
           }
       }

       public override DataTable ConsultarBD(Object objparFamiliaresVO)
       {
           try
           {
               objFamiliaresVO = (Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               if (objFamiliaresVO.getCod() > 0)
               {
                   strSQL.Append(" SELECT");
                   strSQL.Append(" Cod");
                   strSQL.Append(" ,Nome");
                   strSQL.Append(" ,Sexo");
                   strSQL.Append(" ,Idade");
                   strSQL.Append(" ,GanhoTotalMensal");
                   strSQL.Append(" ,GastoTotalMensal");
                   strSQL.Append(" ,Observaçao");
                   strSQL.Append(" FROM");
                   strSQL.Append(" FAMILIARES");
                   strSQL.Append(" WHERE Cod = :parCod");

                   objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                   objComando.Parameters.AddWithValue("parCod", objFamiliaresVO.Cod);

               }
               else if (string.IsNullOrEmpty(objFamiliaresVO.Nome))
               {
                   strSQL.Append(" SELECT");
                   strSQL.Append(" Cod");
                   strSQL.Append(" ,Nome");
                   strSQL.Append(" ,Sexo");
                   strSQL.Append(" ,Idade");
                   strSQL.Append(" ,GanhoTotalMensal");
                   strSQL.Append(" ,GastoTotalMensal");
                   strSQL.Append(" ,Observaçao");
                   strSQL.Append(" FROM");
                   strSQL.Append(" FAMILIARES");

                   objComando = new OleDbCommand(strSQL.ToString(), getConexao());
               }
               else
               {
                   strSQL.Append(" SELECT");
                   strSQL.Append(" Cod");
                   strSQL.Append(" ,Nome");
                   strSQL.Append(" ,Sexo");
                   strSQL.Append(" ,Idade");
                   strSQL.Append(" ,GanhoTotalMensal");
                   strSQL.Append(" ,GastoTotalMensal");
                   strSQL.Append(" ,Observaçao");
                   strSQL.Append(" FROM");
                   strSQL.Append(" FAMILIARES");
                   strSQL.Append(" WHERE Nome = :parNome");

                   objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                   objComando.Parameters.AddWithValue("parNome", objFamiliaresVO.Nome);
               }

               objAdaptador = new OleDbDataAdapter();
               objAdaptador.SelectCommand = objComando;

               objTabela = new DataTable();
               objAdaptador.Fill(objTabela);

               return objTabela;
           }
           catch (Exception ex)
           {

               throw new Exception("Falha ao Consultar Banco De Dados De Familiares " + ex.Message);
           }
       }
       public override void ConsultarBD(ref Object objparFamiliaresVO)
       {
           try
           {
               objFamiliaresVO = (Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               if (objFamiliaresVO.getCod() > 0)
               {
                   strSQL.Append(" SELECT");
                   strSQL.Append(" Cod");
                   strSQL.Append(" ,Nome");
                   strSQL.Append(" ,Sexo");
                   strSQL.Append(" ,Idade");
                   strSQL.Append(" ,GanhoTotalMensal");
                   strSQL.Append(" ,GastoTotalMensal");
                   strSQL.Append(" ,Observaçao");
                   strSQL.Append(" FROM");
                   strSQL.Append(" FAMILIARES");
                   strSQL.Append(" WHERE Cod = :parCod");

                   objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                   objComando.Parameters.AddWithValue("parCod", objFamiliaresVO.Cod);

               }
               else if (string.IsNullOrEmpty(objFamiliaresVO.Nome))
               {
                   strSQL.Append(" SELECT");
                   strSQL.Append(" Cod");
                   strSQL.Append(" ,Nome");
                   strSQL.Append(" ,Sexo");
                   strSQL.Append(" ,Idade");
                   strSQL.Append(" ,GanhoTotalMensal");
                   strSQL.Append(" ,GastoTotalMensal");
                   strSQL.Append(" ,Observaçao");
                   strSQL.Append(" FROM");
                   strSQL.Append(" FAMILIARES");

                   objComando = new OleDbCommand(strSQL.ToString(), getConexao());
               }
               else
               {
                   strSQL.Append(" SELECT");
                   strSQL.Append(" Cod");
                   strSQL.Append(" ,Nome");
                   strSQL.Append(" ,Sexo");
                   strSQL.Append(" ,Idade");
                   strSQL.Append(" ,GanhoTotalMensal");
                   strSQL.Append(" ,GastoTotalMensal");
                   strSQL.Append(" ,Observaçao");
                   strSQL.Append(" FROM");
                   strSQL.Append(" FAMILIARES");
                   strSQL.Append(" WHERE Nome = :parNome");

                   objComando = new OleDbCommand(strSQL.ToString(), getConexao());

                   objComando.Parameters.AddWithValue("parNome", objFamiliaresVO.Nome);
               }

               objAdaptador = new OleDbDataAdapter();
               objAdaptador.SelectCommand = objComando;

               objTabela = new DataTable();
               objAdaptador.Fill(objTabela);

               foreach (DataRow itemFamiliaresVO in objTabela.Rows)
               {
                   Familiares_VO objItemFamiliaresVO = new Familiares_VO(Convert.ToInt32(itemFamiliaresVO["Cod"].ToString()),
                                                       itemFamiliaresVO["Nome"].ToString(),
                                                       itemFamiliaresVO["Sexo"].ToString(),
                                                       Convert.ToInt32(itemFamiliaresVO["Idade"].ToString()),
                                                       Convert.ToDouble(itemFamiliaresVO["GanhoTotalMensal"].ToString()),
                                                       Convert.ToDouble(itemFamiliaresVO["GastoTotalMensal"].ToString()),
                                                       itemFamiliaresVO["Descricao"].ToString());
                   objFamiliaresVO.objFamiliaresVOCollection.Add(objFamiliaresVO);
               }
           }
           catch (Exception ex)
           {

               throw new Exception("Falha ao Consultar Banco De Dados De Familiares Desconectado " + ex.Message);
           }
       }
       public override bool InserirBD(Object objparFamiliaresVO)
       {
           try
           {
               abreConexao();

               bool resultado = false;

               objFamiliaresVO = (Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" INSERT INTO");
               strSQL.Append(" FAMILIARES(");
               strSQL.Append(" Nome");
               strSQL.Append(" ,Sexo");
               strSQL.Append(" ,Idade");
               strSQL.Append(" ,GanhoTotalMensal");
               strSQL.Append(" ,GastoTotalMensal");
               strSQL.Append(" ,Observaçao");
               strSQL.Append(" )VALUES(");
               strSQL.Append(" :parNome");
               strSQL.Append(" ,:parSexo");
               strSQL.Append(" ,:parIdade");
               strSQL.Append(" ,:parGanhoTotalMensal");
               strSQL.Append(" ,:parGastoTotalMensal");
               strSQL.Append(" ,:parObservaçao)");

               objComando = new OleDbCommand(strSQL.ToString(), getConexao());

               objComando.Parameters.AddWithValue("parNome", objFamiliaresVO.Nome);
               objComando.Parameters.AddWithValue("parSexo", objFamiliaresVO.Sexo);
               objComando.Parameters.AddWithValue("parIdade", objFamiliaresVO.Idade);
               objComando.Parameters.AddWithValue("parGanhoTotalMensal", objFamiliaresVO.GanhoTotalMensal);
               objComando.Parameters.AddWithValue("parGastoTotalMensal", objFamiliaresVO.GastoTotalMensal);
               objComando.Parameters.AddWithValue("parObservaçao", objFamiliaresVO.Observaçao);

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

               throw new Exception("Falha ao Inserir no Banco De Dados De Familiares " + ex.Message);
           }
       }
       public override bool ExcluirBD(Object objparFamiliaresVO)
       {
           try
           {
               abreConexao();

               bool resultado = false;

               objFamiliaresVO = (Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" DELETE FROM");
               strSQL.Append(" FAMILIARES");
               strSQL.Append(" WHERE");
               strSQL.Append(" Cod= :parCod");

               objComando = new OleDbCommand(strSQL.ToString(), getConexao());

               objComando.Parameters.AddWithValue("parCod", objFamiliaresVO.Cod);

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

               throw new Exception("Falha ao Excluir no Banco De Dados De Familiares " + ex.Message);
           }
       }

       public override bool AlterarBD(Object objparFamiliaresVO)
       {
           try
           {
               abreConexao();

               bool resultado = false;

               objFamiliaresVO = (Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" UPDATE");
               strSQL.Append(" FAMILIARES");
               strSQL.Append(" SET");
               strSQL.Append(" Nome= :parNome");
               strSQL.Append(" ,Sexo= :parSexo");
               strSQL.Append(" ,Idade= :parIdade");
               strSQL.Append(" ,GanhoTotalMensal= :parGanhoTotalMensal");
               strSQL.Append(" ,GastoTotalMensal= :parGastoTotalMensal");
               strSQL.Append(" ,Observaçao= :parObservaçao");
               strSQL.Append(" WHERE");
               strSQL.Append(" Cod= :parCod");


               objComando = new OleDbCommand(strSQL.ToString(), getConexao());

               objComando.Parameters.AddWithValue("parNome", objFamiliaresVO.Nome);
               objComando.Parameters.AddWithValue("parSexo", objFamiliaresVO.Sexo);
               objComando.Parameters.AddWithValue("parIdade", objFamiliaresVO.Idade);
               objComando.Parameters.AddWithValue("parGanhoTotalMensal", objFamiliaresVO.GanhoTotalMensal);
               objComando.Parameters.AddWithValue("parGastoTotalMensal", objFamiliaresVO.GastoTotalMensal);
               objComando.Parameters.AddWithValue("parObservaçao", objFamiliaresVO.Observaçao);
               objComando.Parameters.AddWithValue("parCod", objFamiliaresVO.Cod);

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

               throw new Exception("Falha ao Alterar no Banco De Dados De Familiares " + ex.Message);
           }
       }
    }
}
