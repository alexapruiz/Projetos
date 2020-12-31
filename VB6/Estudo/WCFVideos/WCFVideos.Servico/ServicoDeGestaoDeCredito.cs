using System;
using WCFVideos.Contratos;

namespace WCFVideos.Servico
{
    public class ServicoDeGestaoDeCredito:IGestorDeCredito
    {
            public decimal RecuperarQuantidadeDeRecursoDisponivel()
            {
                //buscando informações em DB
                return 1000.0M;
            }

            public void AnalisarProposta(Proposta proposta)
            {
                //Enfileirando em algum repositorio
            }

            public void EfetivarProposta(Proposta proposta)
            {
                //Enviar a proposta
            }

            public Proposta[] RecuperarProposta(Status status)
            {
                throw new NotImplementedException();
            }
    }
 }