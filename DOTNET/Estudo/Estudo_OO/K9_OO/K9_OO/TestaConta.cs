using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K9_OO
{
    public class TestaConta
    {
        static void Main ()
        {
            //Conta conta = new Conta();
            //conta.numero = 1192321;
            //conta.saldo = 2500;
            //Console.WriteLine("O Número da Conta é: " + conta.numero);
            //Console.WriteLine("O Saldo da Conta é: " + conta.saldo);
            //Console.WriteLine("O Limite da Conta é: " + conta.limite);
            
            //CartaoDeCredito cdc = new CartaoDeCredito();
            //Cliente cliente = new Cliente ();

            //// Ligando os objetos
            //cdc.cliente = cliente;
            //cdc.numero = 12345678;
            //cdc.cliente.nome = " Rafael Cosentino ";
            //Console.WriteLine("O Número do Cartão é:" + cdc.numero);
            //Console.WriteLine("O Nome do Cliente é:" + cdc.cliente.nome);

            // Referência de um objeto
            //Conta conta = new Conta (123);
            // Chamando o método Deposita ()
            //conta.Deposita(1000);
            //Console.WriteLine("O número da conta é:" + conta.numero);
            //Console.WriteLine("O novo saldo da conta é:" + conta.saldo);

            Conta origem = new Conta(1234);
            Conta destino = new Conta(5678);
            try
            {
                origem.Deposita(1000);
                destino.Deposita(1000);
                origem.Transfere(destino, 1000);
            }
            catch (System.ArgumentException e)
            {
                System.Console.WriteLine("Houve um erro ao depositar ou na transferência");
            }

            Console.WriteLine("O Saldo da conta origem:" + origem.ConsultaSaldo());
            Console.WriteLine("O Saldo da conta destino é:" + destino.ConsultaSaldo());

            //int[] numeros = new int[100];
            //numeros[1] = 1;
            //numeros[2] = 2;
            //for (int i = 0; i < 100; i ++)
            //{
            //    numeros [i] = i;
            //    Console.WriteLine(numeros[i]);
            //}

            Console.WriteLine("Fim");
        }
    }
}