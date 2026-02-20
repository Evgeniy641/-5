using System;

namespace ZooShopLINQ
{
    public class Customer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
        public string Address { get; set; }

        public Customer()
        {
        }

        public Customer(int id, string name, int age, string address)
        {
            Id = id;
            Name = name;
            Age = age;
            Address = address;
        }

        public override string ToString()
        {
            return $"ID: {Id}, Имя: {Name}, Возраст: {Age}, Адрес: {Address}";
        }
    }
}