using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
  
namespace MySelfHostService
{
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string CreditCardNumber { get; set; }
    }
}