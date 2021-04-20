using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model
{
    public class Preferencias_De_Familiares_VO
    {
        Preferencias_VO objPreferenciasVO;
        Familiares_VO objFamiliaresVO;
        float intensidade;
        string observaçao;

        public Preferencias_De_Familiares_VO()
        {

        }

        public Preferencias_De_Familiares_VO(Familiares_VO objFamiliaresVO, Preferencias_VO objPreferenciasVO, float fltIntensidade, string strObservaçao)
        {
            setFamiliaresVO(objFamiliaresVO);
            setPreferenciasVO(objPreferenciasVO);
            setIntensidade(fltIntensidade);
            setObservaçao(strObservaçao);
        }
        public Familiares_VO getFamiliares()
        {
            return this.objFamiliaresVO;
        }
        public void setFamiliaresVO(Familiares_VO objFamiliaresVO)
        {
            this.objFamiliaresVO = objFamiliaresVO;
        }
        public Familiares_VO ObjFamiliarVO
        {
            get { return this.objFamiliaresVO; }
            set { this.objFamiliaresVO = value; }
        }
        public Preferencias_VO getPreferencias()
        {
            return this.objPreferenciasVO;
        }
        public void setPreferenciasVO(Preferencias_VO objPreferenciasVO)
        {
            this.objPreferenciasVO = objPreferenciasVO;
        }
        public Preferencias_VO ObjPreferenciasVO
        {
            get { return this.objPreferenciasVO; }
            set { this.objPreferenciasVO = value; }
        }
        public float  getIntensidade()
        {
            return this.intensidade;
        }
        public void setIntensidade(float fltIntensidade)
        {
            this.intensidade = fltIntensidade;
        }
        public float Intensidade
        {
            get { return this.intensidade; }
            set { this.intensidade = value; }
        }
        public string getObservaçao()
        {
            return this.observaçao;
        }
        public void setObservaçao(string strObservaçao)
        {
            this.observaçao = strObservaçao;
        }
        public string Observaçao
        {
            get { return this.observaçao; }
            set { this.observaçao = value; }
        }
        public List<Preferencias_De_Familiares_VO> objPreferenciasDeFamiliaresVOCollection = new List<Preferencias_De_Familiares_VO>();
    }
}
