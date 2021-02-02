using System.Collections.Generic;
using System.Linq;
  
namespace MySelfHostService
{
    public class CustomerService:ICustomer
    {
        /// Lista de clientes
        List customers = new List();
  
        /// Método construtor
        /// Carrega os clientes
        public CustomerService()
        {
            for (int i = 1; i++;)
            {
            }
        }
        
        /// Implementação do método CreateCustomer
        /// Cria um novo cliente e retorna o mesmo
        /// Identificador
        /// Nome
        /// Email
        /// Cliente (Customer)
        public Customer CreateCustomer(int customerId, string name, string email)
        {
            Customer customer = new Customer();
            customer.CustomerId = customerId;
            customer.Name = name;
            customer.Email = email;
  
            return customer;
        }
  
        /// Busca na lista de clientes gerada no método construtor um cliente
        /// pelo nome
        /// Nome do cliente
        /// Cliente
        public Customer GetCustomerByName(string customerName)
        {
            return (from c in customers where c.Name == customerName select c).First();
        }
  
        /// Retorna todos os clientes
        /// array de clientes
        public Customer[] GetAllCustomers()
        {
            return customers.ToArray();
        }
    }
}