using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Camada_Model;
using System.Data;
using Biblioteca_DAO;

namespace Camada_FD
{
    public class Familiares_FD
    {
        Familiares_DAO objFamiliaresDAO;

        public bool gerarAccess(string strNomeCompletoPlanilha)
        {
            try
            {
                objFamiliaresDAO = new  Familiares_DAO();
                return objFamiliaresDAO.gerarAccess(strNomeCompletoPlanilha);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarBD(Familiares_VO objparFamiliaresVO)
        {
            try
            {
                objFamiliaresDAO = new Familiares_DAO();
                return objFamiliaresDAO.ConsultarBD(objparFamiliaresVO);
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        public void ConsultarBD(ref Familiares_VO objparFamiliaresVO)
        {
            try
            {
                objFamiliaresDAO = new Familiares_DAO();
                objFamiliaresDAO.ConsultarBD(objparFamiliaresVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public bool InserirBD(Familiares_VO objparFamiliaresVO)
        {
            try
            {
                objFamiliaresDAO = new Familiares_DAO();
                return objFamiliaresDAO.InserirBD(objparFamiliaresVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public bool ExcluirBD(Familiares_VO objparFamiliaresVO)
        {
            try
            {
                objFamiliaresDAO = new Familiares_DAO();
                return objFamiliaresDAO.ExcluirBD(objparFamiliaresVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public bool AlterarBD(Familiares_VO objparFamiliaresVO)
        {
            try
            {
                objFamiliaresDAO = new Familiares_DAO();
                return objFamiliaresDAO.AlterarBD(objparFamiliaresVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
