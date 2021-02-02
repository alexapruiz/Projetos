namespace MySelfHostService
{
    public interface ICustomer
    {
        Customer CreateCustomer(int customerId, string name, string email);
        Customer GetCustomerByName(string customerName);
        Customer[] GetAllCustomers();
    }
}