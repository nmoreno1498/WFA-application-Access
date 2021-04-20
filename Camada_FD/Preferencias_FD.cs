using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Camada_Model;
using Biblioteca_DAO;

namespace Camada_FD
{
    public class Preferencias_FD
    {
        Preferencias_DAO objPreferenciasDAO;

        public bool gerarAccess(string strNomeCompletoPlanilha)
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.gerarAccess(strNomeCompletoPlanilha);
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }

        public List<string> ImportaBD()
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.ImportaBD();
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        public List<string> ImportaBDDesconectado()
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.ImportaBDDesconectado();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarBD(Preferencias_VO objPreferenciasVO)
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.ConsultarBD(objPreferenciasVO);
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        public void ConsultarBD(ref Preferencias_VO objPreferenciasVO)
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                objPreferenciasDAO.ConsultarBD(objPreferenciasVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public bool InserirBD(Preferencias_VO objPreferenciasVO)
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.InserirBD(objPreferenciasVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBD(Preferencias_VO objPreferenciasVO)
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.ExcluirBD(objPreferenciasVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public bool AlterarBD(Preferencias_VO objPreferenciasVO)
        {
            try
            {
                objPreferenciasDAO = new Preferencias_DAO();
                return objPreferenciasDAO.AlterarBD(objPreferenciasVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
