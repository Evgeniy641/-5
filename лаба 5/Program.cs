using System;
using System.IO;
using System.Linq;
using System.Text;

namespace ZooShopLINQ
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            ZooShopDataManager dataManager = new ZooShopDataManager();

            Console.WriteLine("ЗООМАГАЗИН - Система управления данными");
            Console.WriteLine("=====================================");

            string defaultPath = @"C:\Users\79223\OneDrive\Рабочий стол\С#\лаба 5\лаба 5\лаба 5\Data\LR5-var2.xls";
            Console.WriteLine($"Попытка автоматической загрузки из файла: {defaultPath}");

            if (File.Exists(defaultPath))
            {
                dataManager.LoadFromExcel(defaultPath);
            }
            else
            {
                Console.WriteLine($"Файл {defaultPath} не найден!");
                Console.WriteLine($"Текущая директория: {Directory.GetCurrentDirectory()}");
                Console.WriteLine("Вы сможете загрузить файл вручную через меню (пункт 1)");
            }

            while (true)
            {
                Console.WriteLine("\n=====================================");
                Console.WriteLine("ГЛАВНОЕ МЕНЮ");
                Console.WriteLine("=====================================");
                Console.WriteLine("1. Загрузить данные из Excel файла (вручную)");
                Console.WriteLine("2. Просмотреть все данные");
                Console.WriteLine("3. Удалить элемент");
                Console.WriteLine("4. Добавить элемент");
                Console.WriteLine("5. Выполнить запросы LINQ");
                Console.WriteLine("6. Сохранить изменения");
                Console.WriteLine("0. Выход");
                Console.Write("Выберите действие: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        Console.Write("Введите путь к файлу (или нажмите Enter для Data/LR5-var2.xls): ");
                        string path = Console.ReadLine();

                        if (string.IsNullOrWhiteSpace(path))
                        {
                            path = "Data/LR5-var2.xls";
                        }

                        dataManager.LoadFromExcel(path);
                        Console.WriteLine("Нажмите любую клавишу...");
                        Console.ReadKey();
                        break;

                    case "2":
                        dataManager.ViewAllData();
                        Console.WriteLine("\nНажмите любую клавишу...");
                        Console.ReadKey();
                        break;

                    case "3":
                        DeleteMenu(dataManager);
                        break;

                    case "4":
                        AddMenu(dataManager);
                        break;

                    case "5":
                        QueryMenu(dataManager);
                        break;

                    case "6":
                        dataManager.SaveToExcel();
                        Console.WriteLine("Нажмите любую клавишу...");
                        Console.ReadKey();
                        break;

                    case "0":
                        Console.WriteLine("Программа завершена.");
                        return;

                    default:
                        Console.WriteLine("Неверный выбор!");
                        Console.ReadKey();
                        break;
                }
            }
        }

        private static void DeleteMenu(ZooShopDataManager dm)
        {
            Console.Clear();
            Console.WriteLine("=== УДАЛЕНИЕ ЭЛЕМЕНТОВ ===");
            Console.WriteLine("1. Удалить животное");
            Console.WriteLine("2. Удалить покупателя");
            Console.WriteLine("3. Удалить продажу");
            Console.Write("Выберите тип: ");

            string choice = Console.ReadLine();
            Console.Write("Введите ID: ");

            if (int.TryParse(Console.ReadLine(), out int id))
            {
                switch (choice)
                {
                    case "1":
                        dm.DeleteAnimal(id);
                        break;

                    case "2":
                        dm.DeleteCustomer(id);
                        break;

                    case "3":
                        dm.DeleteSale(id);
                        break;

                    default:
                        Console.WriteLine("Неверный выбор!");
                        break;
                }
            }
            else
            {
                Console.WriteLine("Некорректный ID!");
            }

            Console.WriteLine("Нажмите любую клавишу...");
            Console.ReadKey();
        }

        private static void AddMenu(ZooShopDataManager dm)
        {
            Console.Clear();
            Console.WriteLine("=== ДОБАВЛЕНИЕ ЭЛЕМЕНТОВ ===");
            Console.WriteLine("1. Добавить животное");
            Console.WriteLine("2. Добавить покупателя");
            Console.WriteLine("3. Добавить продажу");
            Console.Write("Выберите тип: ");

            string choice = Console.ReadLine();

            try
            {
                switch (choice)
                {
                    case "1":
                        Animal animal = new Animal();
                        Console.Write("ID: ");
                        animal.Id = int.Parse(Console.ReadLine());
                        Console.Write("Вид: ");
                        animal.Species = Console.ReadLine();
                        Console.Write("Порода: ");
                        animal.Breed = Console.ReadLine();
                        dm.AddAnimal(animal);
                        break;

                    case "2":
                        Customer customer = new Customer();
                        Console.Write("ID: ");
                        customer.Id = int.Parse(Console.ReadLine());
                        Console.Write("Имя: ");
                        customer.Name = Console.ReadLine();
                        Console.Write("Возраст: ");
                        customer.Age = int.Parse(Console.ReadLine());
                        Console.Write("Адрес: ");
                        customer.Address = Console.ReadLine();
                        dm.AddCustomer(customer);
                        break;

                    case "3":
                        Sale sale = new Sale();
                        Console.Write("ID: ");
                        sale.Id = int.Parse(Console.ReadLine());
                        Console.Write("ID животного: ");
                        sale.AnimalId = int.Parse(Console.ReadLine());
                        Console.Write("ID покупателя: ");
                        sale.CustomerId = int.Parse(Console.ReadLine());
                        Console.Write("Дата (ГГГГ-ММ-ДД): ");
                        sale.Date = DateTime.Parse(Console.ReadLine());
                        Console.Write("Цена: ");
                        sale.Price = decimal.Parse(Console.ReadLine());
                        dm.AddSale(sale);
                        break;

                    default:
                        Console.WriteLine("Неверный выбор!");
                        break;
                }
            }
            catch (FormatException)
            {
                Console.WriteLine("Ошибка: Неверный формат данных!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при добавлении: {ex.Message}");
            }

            Console.WriteLine("Нажмите любую клавишу...");
            Console.ReadKey();
        }

        private static void QueryMenu(ZooShopDataManager dm)
        {
            Console.Clear();
            Console.WriteLine("=== LINQ ЗАПРОСЫ ===");

            if (!dm.Animals.Any())
            {
                Console.WriteLine("Данные не загружены! Сначала загрузите файл (пункт 1).");
                Console.WriteLine("Нажмите любую клавишу...");
                Console.ReadKey();

                return;
            }

            Console.WriteLine("\n1. Кошки породы 'Сфинкс':");
            var sphynxCats = dm.GetSphynxCats();

            if (sphynxCats.Any())
            {
                sphynxCats.ForEach(c => Console.WriteLine($"   {c}"));
            }
            else
            {
                Console.WriteLine("   Не найдено");
            }

            Console.WriteLine($"\n2. Общая сумма продаж покупателям старше 30 лет: {dm.GetTotalSalesForCustomersOver30():C}");

            Console.WriteLine("\n3. Продажи кошек с деталями:");
            var catSales = dm.GetCatSalesWithDetails();

            if (catSales.Any())
            {
                catSales.ForEach(s => Console.WriteLine($"   {s}"));
            }
            else
            {
                Console.WriteLine("   Не найдено");
            }

            Console.WriteLine($"\n4. Средний возраст покупателей собак породы 'Такса': {dm.GetAverageAgeOfDachshundBuyers():F1} лет");

            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }
    }
}