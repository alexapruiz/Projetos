using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K9_OO
{
    public class Cliente
    {
        public string nome;
        public string endereco;
        public double salario;

        //Método construtor da classe
        public Cliente(string nome)
        {
            this.nome = nome;
        }
    }
}
