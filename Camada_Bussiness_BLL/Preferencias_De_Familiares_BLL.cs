using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Camada_FD;
using Camada_Model;

namespace Camada_Bussiness_BLL
{
   public class Preferencias_De_Familiares_BLL
    {
       Preferencias_De_Familiares_FD objPreferenciasDeFamiliaresFD;

       public bool gerarAccess(string strNomeCompletoPlanilha)
       {
           try
           {
               objPreferenciasDeFamiliaresFD = new Preferencias_De_Familiares_FD();
               return objPreferenciasDeFamiliaresFD.gerarAccess(strNomeCompletoPlanilha);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }

       public DataTable ConsultarBD(Object objparPrefFamVO)
       {
           try
           {
               objPreferenciasDeFamiliaresFD = new Preferencias_De_Familiares_FD();
               return objPreferenciasDeFamiliaresFD.ConsultarBD(objparPrefFamVO);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }
       public void ConsultarBD(ref Object objparPrefFamVO)
       {
           try
           {
               objPreferenciasDeFamiliaresFD = new Preferencias_De_Familiares_FD();
               objPreferenciasDeFamiliaresFD.ConsultarBD(objparPrefFamVO);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }
       public bool InserirBD(Object objparPrefFamVO)
       {
           try
           {
               objPreferenciasDeFamiliaresFD = new Preferencias_De_Familiares_FD();
               return objPreferenciasDeFamiliaresFD.InserirBD(objparPrefFamVO);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }
       public bool ExcluirBD(Object objparPrefFamVO)
       {
           try
           {
               objPreferenciasDeFamiliaresFD = new Preferencias_De_Familiares_FD();
               return objPreferenciasDeFamiliaresFD.ExcluirBD(objparPrefFamVO);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }
       public bool AlterarBD(Object objparPrefFamVO)
       {
           try
           {
               objPreferenciasDeFamiliaresFD = new Preferencias_De_Familiares_FD();
               return objPreferenciasDeFamiliaresFD.AlterarBD(objparPrefFamVO);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }
    }
}
