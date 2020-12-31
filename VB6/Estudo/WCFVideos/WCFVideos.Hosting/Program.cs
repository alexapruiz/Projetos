using System;
using System.ServiceModel;
using WCFVideos.Servico;
using WCFVideos.Contratos;
using System.ServiceModel.Description;

namespace WCFVideos.Hosting
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Configuração Imperativa
            using (ServiceHost host =
                new ServiceHost(typeof(ServicoDeGestaoDeCredito), new Uri[] { new Uri("net.tcp://localhost:9876"), new Uri("http://localhost:8766") }))
            {
                host.Description.Behaviors.Add(new ServiceMetadataBehavior() { HttpGetEnabled = true });

                host.AddServiceEndpoint(typeof(IGestorDeCredito), new NetTcpBinding(), "srv");
                host.AddServiceEndpoint(typeof(IGestorDeCredito), new BasicHttpBinding(), "srv");

                //WSDL
                host.AddServiceEndpoint(typeof(IMetadataExchange), MetadataExchangeBindings.CreateMexHttpBinding(), "mex");

                host.Open();
                Console.ReadLine();
            }
            #endregion
        }
    }
}