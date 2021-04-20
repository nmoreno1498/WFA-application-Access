using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Configuration;

namespace Biblioteca_DAO
{
    public  class Sinlgeton_DB_DAO
    {
        private static OleDbConnection objConexao;

        public static OleDbConnection getConexao()
        {
            if (objConexao == null)
            {
                objConexao = new OleDbConnection(ConfigurationSettings.AppSettings["connectionstring"].ToString());
            }
            return objConexao;
        }
        public static void abreConexao()
        {
            if (getConexao().State == ConnectionState.Closed)
            {
                objConexao.Open();
            }
        }
        public static void fechaConexao()
        {
            if (getConexao().State == ConnectionState.Open)
            {
                objConexao.Close();

                objConexao.Dispose();

                objConexao = null;
            }
        }
    }
}
