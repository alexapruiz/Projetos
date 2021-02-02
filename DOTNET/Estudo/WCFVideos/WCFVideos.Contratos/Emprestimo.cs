using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WCFVideos.Contratos
{
    public class Emprestimo
    {
        public decimal Valor { get; set; }
        public int QuantidadeDeParcelas { get; set; }
        public decimal TaxaDeJuros { get; set; }
    }
}
