using System;

namespace ZooShopLINQ
{
    public class Sale
    {
        public int Id { get; set; }
        public int AnimalId { get; set; }
        public int CustomerId { get; set; }
        public DateTime Date { get; set; }
        public decimal Price { get; set; }

        public Sale()
        {
        }

        public Sale(int id, int animalId, int customerId, DateTime date, decimal price)
        {
            Id = id;
            AnimalId = animalId;
            CustomerId = customerId;
            Date = date;
            Price = price;
        }

        public override string ToString()
        {
            return $"ID: {Id}, Животное: {AnimalId}, Покупатель: {CustomerId}, Дата: {Date.ToShortDateString()}, Цена: {Price}";
        }
    }
}