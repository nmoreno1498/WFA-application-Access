using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Biblioteca_DAO;
using System.Data;
using System.Data.OleDb;
using Camada_Model;

namespace Camada_FD
{
   public  class Preferencias_De_Familiares_FD
    {
       Preferencias_De_Familiares_DAO objPreferenciasDeFamiliaresDAO;

       public bool gerarAccess(string strNomeCompletoPlanilha)
       {
           try
           {
               objPreferenciasDeFamiliaresDAO = new Preferencias_De_Familiares_DAO();
               return objPreferenciasDeFamiliaresDAO.gerarAccess(strNomeCompletoPlanilha);
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
               objPreferenciasDeFamiliaresDAO = new Preferencias_De_Familiares_DAO();
               return objPreferenciasDeFamiliaresDAO.ConsultarBD(objparPrefFamVO);
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
               objPreferenciasDeFamiliaresDAO = new Preferencias_De_Familiares_DAO();
               objPreferenciasDeFamiliaresDAO.ConsultarBD(objparPrefFamVO);
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
               objPreferenciasDeFamiliaresDAO = new Preferencias_De_Familiares_DAO();
               return objPreferenciasDeFamiliaresDAO.InserirBD(objparPrefFamVO);
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
               objPreferenciasDeFamiliaresDAO = new Preferencias_De_Familiares_DAO();
               return objPreferenciasDeFamiliaresDAO.ExcluirBD(objparPrefFamVO);
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
               objPreferenciasDeFamiliaresDAO = new Preferencias_De_Familiares_DAO();
               return objPreferenciasDeFamiliaresDAO.AlterarBD(objparPrefFamVO);
           }
           catch (Exception ex)
           {

               throw ex;
           }
       }
    }
}
