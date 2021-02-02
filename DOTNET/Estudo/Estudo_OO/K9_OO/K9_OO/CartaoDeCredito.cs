using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K9_OO
{
    public class CartaoDeCredito
    {
        public int numero;
        public DateTime dataDeValidade;
        public Cliente cliente;

        //Método construtor da classe
        public CartaoDeCredito(int numero)
        {
            this.numero = numero;
        }
    }
}
