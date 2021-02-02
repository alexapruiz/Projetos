using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K9_OO
{
class Servico
    {
        public Cliente Contratante { get ; set ; }
        public Funcionario Responsavel { get ; set ; }
        
        //GERAL
        public string DataDeContratacao { get ; set ; }
        public double Valor { get ; set ; }
        public double Taxa { get ; set ; }

        // SEGURO DE VEICULO
        public Veiculo Veiculo { get ; set ; }
        public double ValorDoSeguroDeVeiculo { get ; set ; }
        public double Franquia { get ; set ; }
    }
}