using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Camada_Model;
using Camada_FD;
using System.IO;

namespace Camada_Bussiness_BLL
{
    public class Preferencias_BLL
    {
        Preferencias_FD objPreferenciasFD;
        string strLinhaLida;
        StreamReader objLeitor;

        public bool gerarAccess(string strNomeCompletoPlanilha)
        {
            try
            {
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.gerarAccess(strNomeCompletoPlanilha);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public List<string> ImportaTextoWhile()
        {
            try
            {
                List<string> resultado = new List<string>();

                objLeitor = new StreamReader(@"C:\Users\nmore\OneDrive\Documentos\cualquierCosa.txt");
                strLinhaLida = objLeitor.ReadLine();

                while (strLinhaLida != null)
                {
                    resultado.Add(strLinhaLida);
                    strLinhaLida = objLeitor.ReadLine();
                }
                objLeitor.Close();

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha ao Importar Texto While" + ex.Message);
            }
        }

        public List<string> ImportaBD()
        {
            try
            {
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.ImportaBD();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public List<string> ImportaBDDseconectado()
        {
            try
            {
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.ImportaBDDesconectado();
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
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.ConsultarBD(objPreferenciasVO);
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
                objPreferenciasFD = new Preferencias_FD();
                objPreferenciasFD.ConsultarBD(objPreferenciasVO);
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
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.InserirBD(objPreferenciasVO);
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
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.ExcluirBD(objPreferenciasVO);
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
                objPreferenciasFD = new Preferencias_FD();
                return objPreferenciasFD.AlterarBD(objPreferenciasVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
