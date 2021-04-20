using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model
{
   public class Familiares_VO
    {
       private int cod;

       private string nome;

       private string sexo;

       private int idade;

       private double ganhototalmensal;

       private double gastototalmensal;

       private string observaçao;


       public Familiares_VO()
       {

       }
       public Familiares_VO(int intCod, string strNome, string strSexo)
       {
           setCod(intCod);
           setNome(strNome);
           setSexo(strSexo);
       }
       public Familiares_VO(int intCod, string strNome, string strSexo, int intIdade, double dbGanhoTotalMensal, double dbGastoTotalMensal, string strObservaçao)
       {
           setCod(intCod);
           setNome(strNome);
           setSexo(strSexo);
           setIdade(intIdade);
           setGanhoTotalMensal(dbGanhoTotalMensal);
           setGastoTotalMensal(dbGastoTotalMensal);
           setObservaçao(strObservaçao);
       }
       public int getCod()
       {
           return this.cod;
       }
       public void setCod(int intCod)
       {
           this.cod = intCod;
       }
       public int Cod
       {
           get { return this.cod; }
           set { this.cod = value; }
       }

       public string getNome()
       {
           return this.nome;
       }
       public void setNome(string strNome)
       {
           this.nome = strNome;
       }
       public string Nome
       {
           get { return this.nome; }
           set { this.nome = value; }
       }

       public string getSexo()
       {
           return this.sexo;
       }
       public void setSexo(string strSexo)
       {
           if (strSexo == "MASCULINO" || strSexo == "FEMENINO" || strSexo == "INDEFINIDO")
           {
               this.sexo = strSexo;
           }
       }
       public string Sexo
       {
           get { return this.sexo; }
           set
           {
               if (value == "MASCULINO" || value == "FEMENINO" || value == "INDEFINIDO")
               {
                   this.sexo = value;
               }
               else
               {
                   throw new Exception("Atributo Sexo Inexistente ");
               }
           }
       }

       public int getIdade()
       {
           return this.idade;
       }
       public void setIdade(int intIdade)
       {
           this.idade = intIdade;
       }
       public int Idade
       {
           get { return this.idade; }
           set { this.idade = value; }
       }

       public double getGanhoTotalMensal()
       {
           return this.ganhototalmensal;
       }
       public void setGanhoTotalMensal(double dbGanhoTotalMensal)
       {
           this.ganhototalmensal = dbGanhoTotalMensal;
       }
       public double GanhoTotalMensal
       {
           get { return this.ganhototalmensal; }
           set { this.ganhototalmensal = value; }
       }

       public double getGastoTotalMensal()
       {
           return this.gastototalmensal;
       }
       public void setGastoTotalMensal(double dbGastoTotalMensal)
       {
           this.gastototalmensal = dbGastoTotalMensal;
       }
       public double GastoTotalMensal
       {
           get { return this.gastototalmensal; }
           set { this.gastototalmensal = value; }
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
       public List<Familiares_VO> objFamiliaresVOCollection = new List<Familiares_VO>();
    }
}
