using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace WCFVideos.Contratos
{
    [ServiceContract]
    public interface IGestorDeCredito
    {
        [OperationContract]
        decimal RecuperarQuantidadeDeRecursoDisponivel();

        [OperationContract]
        void AnalisarProposta(Proposta proposta);
        
        [OperationContract]
        void EfetivarProposta(Proposta proposta);

        [OperationContract]
        Proposta[] RecuperarProposta(Status status);
    }
}
