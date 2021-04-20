using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Camada_Model;

namespace Biblioteca_DAO
{
   public class Preferencias_De_Familiares_DAO : DAO_DAO
    {
       Preferencias_De_Familiares_VO objPreferenciasDeFamiliaresVO;
       OleDbCommand objComando;
       OleDbDataAdapter objAdaptador;
       DataTable objTabela;

       public bool gerarAccess(string strNomeCompletoDaPlanilha)
       {
           try
           {
               bool resultado = false;
               abreConexao();

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" SELECT");
               strSQL.Append(" Cod");
               strSQL.Append(" ,ID ");
               strSQL.Append(" ,Intensidade");
               strSQL.Append(" ,Observaçao");
               strSQL.Append(" INTO");
               strSQL.Append(" [Excel 8.0;DATABASE=" + strNomeCompletoDaPlanilha + "].[Exportar Excel]");
               strSQL.Append(" FROM");
               strSQL.Append(" Preferencias_De_Familiares");

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
               objPreferenciasDeFamiliaresVO = (Preferencias_De_Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" SELECT");
               strSQL.Append(" Cod");
               strSQL.Append(" ,ID ");
               strSQL.Append(" ,Intensidade");
               strSQL.Append(" ,Observaçao");
               strSQL.Append(" FROM");
               strSQL.Append(" Preferencias_De_Familiares");

               objComando = new OleDbCommand();
               objComando.Connection = getConexao();

               if (objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod > 0 || objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID> 0)
               {
                   if (objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod > 0)
                   {
                       if (objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID> 0)
                       {
                           strSQL.Append(" WHERE");
                           strSQL.Append(" Cod = :parCod");
                           strSQL.Append(" AND ID = :parID");

                           objComando.Parameters.AddWithValue("parCod",objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
                           objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);
                       }
                       else
                       {
                           strSQL.Append(" WHERE");
                           strSQL.Append(" Cod = :parCod");

                           objComando.Parameters.AddWithValue("parCod", objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
                       }
                   }
                   else
                   {
                       if (objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID > 0)
                       {
                           strSQL.Append(" WHERE");
                           strSQL.Append(" ID = :parID");

                           objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);
                       }
                   }
               }
               objComando.CommandText = strSQL.ToString();

               objAdaptador = new OleDbDataAdapter();
               objAdaptador.SelectCommand = objComando;

               objTabela = new DataTable();
               objAdaptador.Fill(objTabela);

               return objTabela;
           }
           catch (Exception ex)
           {
               throw new Exception("Falha ao Consultar BD Preferencia De Familiar ==>" + ex.Message);
           }
       }
       public override void ConsultarBD(ref Object objparFamiliaresVO)
       {
           try
           {
               objPreferenciasDeFamiliaresVO = (Preferencias_De_Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" SELECT");
               strSQL.Append(" Cod");
               strSQL.Append(" ,ID ");
               strSQL.Append(" ,Intensidade");
               strSQL.Append(" ,Observaçao");
               strSQL.Append(" FROM");
               strSQL.Append(" Preferencias_De_Familiares");

               objComando = new OleDbCommand();
               objComando.Connection = getConexao();

               if (objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod > 0 || objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID > 0)
               {
                   if (objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod > 0)
                   {
                       if (objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID > 0)
                       {
                           strSQL.Append(" WHERE");
                           strSQL.Append(" Cod = :parCod");
                           strSQL.Append(" AND ID = :parID");

                           objComando.Parameters.AddWithValue("parCod", objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
                           objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);
                       }
                       else
                       {
                           strSQL.Append(" WHERE");
                           strSQL.Append(" Cod = :parCod");

                           objComando.Parameters.AddWithValue("parCod", objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
                       }
                   }
                   else
                   {
                       if (objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID > 0)
                       {
                           strSQL.Append(" WHERE");
                           strSQL.Append(" ID = :parID");

                           objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);
                       }
                   }
               }
               objComando.CommandText = strSQL.ToString();

               objAdaptador = new OleDbDataAdapter();
               objAdaptador.SelectCommand = objComando;

               objTabela = new DataTable();
               objAdaptador.Fill(objTabela);

               foreach (DataRow itemPrefFam in objTabela.Rows)
               {
                   Preferencias_De_Familiares_VO objItemPreferenciasDeFamiliaresVO = new Preferencias_De_Familiares_VO();

                   objPreferenciasDeFamiliaresVO.ObjFamiliarVO = new Familiares_VO();
                   objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod = Convert.ToInt32(itemPrefFam["Cod"].ToString());

                   Familiares_DAO objFamiliarDAO = new Familiares_DAO();
                   Object objparObjectFamiliar = (Object)objPreferenciasDeFamiliaresVO.ObjFamiliarVO;
                   objFamiliarDAO.ConsultarBD(ref objparObjectFamiliar);

                   objPreferenciasDeFamiliaresVO.ObjFamiliarVO = objPreferenciasDeFamiliaresVO.ObjFamiliarVO.objFamiliaresVOCollection.First<Familiares_VO>();

                   objPreferenciasDeFamiliaresVO.ObjPreferenciasVO = new Preferencias_VO();

                   Preferencias_DAO objPreferenciaDAO = new Preferencias_DAO();
                   Object objparObjectPreferencia = (Object)objPreferenciasDeFamiliaresVO.ObjPreferenciasVO;
                   objPreferenciaDAO.ConsultarBD(ref objparObjectPreferencia);

                   objPreferenciasDeFamiliaresVO.ObjPreferenciasVO = objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.objPreferenciasVOCollection.First<Preferencias_VO>();

                   objPreferenciasDeFamiliaresVO.ObjFamiliarVO = objPreferenciasDeFamiliaresVO.ObjFamiliarVO.objFamiliaresVOCollection.First<Familiares_VO>();

                   objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID = Convert.ToInt32(itemPrefFam["ID"].ToString());

                   objPreferenciasDeFamiliaresVO.Intensidade = Convert.ToSingle(itemPrefFam["Intensidade"].ToString());
                   objPreferenciasDeFamiliaresVO.Observaçao = itemPrefFam["Observaçao"].ToString();

                   objPreferenciasDeFamiliaresVO.objPreferenciasDeFamiliaresVOCollection.Add(objItemPreferenciasDeFamiliaresVO);
               }
           }
           catch (Exception ex)
           {
               throw new Exception("Falha ao Consultar BD Preferencia De Familiar Referenciado ==>" + ex.Message);
           }
       }
       public override bool InserirBD(Object objparFamiliaresVO)
       {
           try
           {
               abreConexao();

               bool resultado = false;

               objPreferenciasDeFamiliaresVO = (Preferencias_De_Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" INSERT INTO");
               strSQL.Append(" Preferencias_De_Familiares(");
               strSQL.Append(" Cod");
               strSQL.Append(" ,ID ");
               strSQL.Append(" ,Intensidade");
               strSQL.Append(" ,Observaçao");
               strSQL.Append(" )VALUES(");
               strSQL.Append(" :parCod");
               strSQL.Append(" ,:parID");
               strSQL.Append(" ,:parIntensidade");
               strSQL.Append(" ,:parObservaçao)");

               objComando = new OleDbCommand(strSQL.ToString(), getConexao());


               objComando.Parameters.AddWithValue("parCod", objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
               objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);
               objComando.Parameters.AddWithValue("parIntensidade", objPreferenciasDeFamiliaresVO.Intensidade);
               objComando.Parameters.AddWithValue("parObservaçao", objPreferenciasDeFamiliaresVO.Observaçao);

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
               throw new Exception("Falha ao Consultar BD Preferencia De Familiar Referenciado ==>" + ex.Message);
           }
           finally
           {
               fechaConexao();
           }
       }
       public override bool ExcluirBD(Object objparFamiliaresVO)
       {
           try
           {
               abreConexao();

               bool resultado = false;

               objPreferenciasDeFamiliaresVO = (Preferencias_De_Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" DELETE FROM");
               strSQL.Append(" Preferencias_De_Familiares");
               strSQL.Append(" WHERE Cod = :parCod");
               strSQL.Append(" AND ID = :parID ");

               objComando = new OleDbCommand(strSQL.ToString(), getConexao());

               objComando.Parameters.AddWithValue("parCod", objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
               objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);

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
               throw new Exception("Falha ao Excluir BD Preferencia De Familiar Referenciado ==>" + ex.Message);
           }
           finally
           {
               fechaConexao();
           }
       }
       public override bool AlterarBD(Object objparFamiliaresVO)
       {
           try
           {
               abreConexao();

               bool resultado = false;

               objPreferenciasDeFamiliaresVO = (Preferencias_De_Familiares_VO)objparFamiliaresVO;

               StringBuilder strSQL = new StringBuilder();

               strSQL.Append(" UPDATE");
               strSQL.Append(" Preferencias_De_Familiares");
               strSQL.Append(" SET");
               strSQL.Append(" Intensidade = :parIntensidade");
               strSQL.Append(" ,Observaçao = :parObservaçao");
               strSQL.Append(" WHERE");
               strSQL.Append(" Cod = :parCod");
               strSQL.Append(" AND");
               strSQL.Append(" ID = :parID");

               objComando = new OleDbCommand(strSQL.ToString(), getConexao());

               objComando.Parameters.AddWithValue("parIntensidade", objPreferenciasDeFamiliaresVO.Intensidade);
               objComando.Parameters.AddWithValue("parObservaçao", objPreferenciasDeFamiliaresVO.Observaçao);
               objComando.Parameters.AddWithValue("parCod", objPreferenciasDeFamiliaresVO.ObjFamiliarVO.Cod);
               objComando.Parameters.AddWithValue("parID", objPreferenciasDeFamiliaresVO.ObjPreferenciasVO.ID);

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
               throw new Exception("Falha ao Alterar BD Preferencia De Familiar ==>" + ex.Message);
           }
           finally
           {
               fechaConexao();
           }
       }
    }
}
