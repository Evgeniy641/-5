using System;

namespace ZooShopLINQ
{
    public class Animal
    {
        public int Id { get; set; }
        public string Species { get; set; }
        public string Breed { get; set; }

        public Animal()
        {
        }

        public Animal(int id, string species, string breed)
        {
            Id = id;
            Species = species;
            Breed = breed;
        }

        public override string ToString()
        {
            return $"ID: {Id}, Вид: {Species}, Порода: {Breed}";
        }
    }
}