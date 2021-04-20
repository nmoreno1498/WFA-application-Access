using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model
{
    public class Preferencias_VO
    {
        private string descricao;

        private int id;


        public Preferencias_VO()
        {

        }
        public Preferencias_VO(string strDescricao)
        {
            setDescricao(strDescricao);
        }
        public Preferencias_VO(int intiD, string strDescricao)
        {
            setiD(intiD);
            setDescricao(strDescricao);
        }
        public int getiD()
        {
            return this.id;
        }
        public void setiD(int intiD)
        {
            this.id = intiD;
        }
        public int ID
        {
            get { return this.id; }
            set { this.id = value; }
        }
        public string getDescricao()
        {
            return this.descricao;
        }
        public void setDescricao(string strDescricao)
        {
            this.descricao = strDescricao;
        }
        public string Descricao
        {
            get { return this.descricao; }
            set { this.descricao = value; }
        }
        public List<Preferencias_VO> objPreferenciasVOCollection = new List<Preferencias_VO>();
    }
}
