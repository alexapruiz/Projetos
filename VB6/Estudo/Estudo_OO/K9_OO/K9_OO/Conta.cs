using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K9_OO
{
    public class Conta
    {
        double saldo;
        double limite = 1000;
        int numero;

        //Método construtor da classe
        public Conta(int numero)
        {
            this.numero = numero;
        }

        public void Deposita ( double valor )
        {
            this.saldo += valor ;
        }

        public double ConsultaSaldo()
        {
            return this.saldo + this.limite;
        }

        public void Transfere (Conta destino, double valor)
        {
            this.saldo -= valor ;
            destino.saldo += valor ;
        }
    }
}