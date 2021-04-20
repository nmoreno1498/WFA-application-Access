using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_Model;
using Camada_FD;

namespace Camada_Bussiness_BLL
{
    public class Familiares_BLL
    {
        Familiares_FD objFamiliaresFD;

        public bool gerarAccess(string strNomeCompletoPlanilha)
        {
            try
            {
                objFamiliaresFD = new Familiares_FD();
                return objFamiliaresFD.gerarAccess(strNomeCompletoPlanilha);
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
                objFamiliaresFD = new Familiares_FD();
                return objFamiliaresFD.ConsultarBD(objparFamiliaresVO);
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
                objFamiliaresFD = new Familiares_FD();
                objFamiliaresFD.ConsultarBD(objparFamiliaresVO);
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
                objFamiliaresFD = new Familiares_FD();
                return objFamiliaresFD.InserirBD(objparFamiliaresVO);
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
                objFamiliaresFD = new Familiares_FD();
                return objFamiliaresFD.ExcluirBD(objparFamiliaresVO);
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
                objFamiliaresFD = new Familiares_FD();
                return objFamiliaresFD.AlterarBD(objparFamiliaresVO);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
