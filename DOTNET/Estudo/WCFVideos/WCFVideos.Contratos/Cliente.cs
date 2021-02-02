using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;

namespace WCFVideos.Contratos
{
    [DataContract]
    public class Cliente
    {
        [DataMember]
        public string Nome {get; set;}

        [DataMember]
        public int Idade {get; set;}

        [DataMember]
        public string Empresa {get; set;}

        [DataMember]
        public decimal Salario {get; set;}
    }
}
