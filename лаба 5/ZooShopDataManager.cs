using ExcelDataReader;
using ExcelDataReader.Exceptions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ZooShopLINQ
{
    public class ZooShopDataManager
    {
        public List<Animal> Animals { get; private set; }
        public List<Customer> Customers { get; private set; }
        public List<Sale> Sales { get; private set; }
        public string FilePath { get; private set; }

        public ZooShopDataManager()
        {
            Animals = new List<Animal>();
            Customers = new List<Customer>();
            Sales = new List<Sale>();
        }

        public void LoadFromExcel(string filePath = @"C:\Users\79223\OneDrive\Рабочий стол\С#\лаба 5\лаба 5\лаба 5\Data\LR5-var2.xls")
        {
            try
            {
                FilePath = filePath;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Файл не найден по пути: {Path.GetFullPath(filePath)}");
                    Console.WriteLine("Попробуйте указать полный путь к файлу.");

                    return;
                }

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        });

                        var animalsTable = result.Tables["Животные"];
                        var customersTable = result.Tables["Покупатели"];
                        var salesTable = result.Tables["Продажи"];

                        Animals.Clear();

                        for (int i = 0; i < animalsTable.Rows.Count; i++)
                        {
                            var row = animalsTable.Rows[i];

                            Animals.Add(new Animal
                            {
                                Id = Convert.ToInt32(row["ID"]),
                                Species = row["Вид"].ToString(),
                                Breed = row["Порода"].ToString()
                            });
                        }

                        Customers.Clear();

                        for (int i = 0; i < customersTable.Rows.Count; i++)
                        {
                            var row = customersTable.Rows[i];

                            Customers.Add(new Customer
                            {
                                Id = Convert.ToInt32(row["ID"]),
                                Name = row["Имя"].ToString(),
                                Age = Convert.ToInt32(row["Возраст"]),
                                Address = row["Адрес"].ToString()
                            });
                        }

                        Sales.Clear();

                        for (int i = 0; i < salesTable.Rows.Count; i++)
                        {
                            var row = salesTable.Rows[i];

                            Sales.Add(new Sale
                            {
                                Id = Convert.ToInt32(row["ID"]),
                                AnimalId = Convert.ToInt32(row["ID животного"]),
                                CustomerId = Convert.ToInt32(row["ID покупателя"]),
                                Date = Convert.ToDateTime(row["Дата"]),
                                Price = Convert.ToDecimal(row["Цена"])
                            });
                        }

                        Console.WriteLine("Данные успешно загружены!");
                        Console.WriteLine($"Загружено: {Animals.Count} животных, {Customers.Count} покупателей, {Sales.Count} продаж");
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Ошибка: Файл не найден!");
            }
            catch (ExcelReaderException ex)
            {
                Console.WriteLine($"Ошибка чтения Excel: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке данных: {ex.Message}");
            }
        }

        public void ViewAllData()
        {
            Console.WriteLine("\n=== ЖИВОТНЫЕ ===");

            foreach (var animal in Animals)
            {
                Console.WriteLine(animal);
            }

            Console.WriteLine("\n=== ПОКУПАТЕЛИ ===");

            foreach (var customer in Customers)
            {
                Console.WriteLine(customer);
            }

            Console.WriteLine("\n=== ПРОДАЖИ ===");

            foreach (var sale in Sales)
            {
                Console.WriteLine(sale);
            }
        }

        public bool DeleteAnimal(int id)
        {
            var animal = Animals.FirstOrDefault(a => a.Id == id);

            if (animal != null)
            {
                if (Sales.Any(s => s.AnimalId == id))
                {
                    Console.WriteLine("Нельзя удалить животное, т.к. есть связанные продажи!");

                    return false;
                }

                Animals.Remove(animal);
                Console.WriteLine($"Животное с ID {id} удалено");

                return true;
            }

            Console.WriteLine($"Животное с ID {id} не найдено");

            return false;
        }

        public bool DeleteCustomer(int id)
        {
            var customer = Customers.FirstOrDefault(c => c.Id == id);

            if (customer != null)
            {
                if (Sales.Any(s => s.CustomerId == id))
                {
                    Console.WriteLine("Нельзя удалить покупателя, т.к. есть связанные продажи!");

                    return false;
                }

                Customers.Remove(customer);
                Console.WriteLine($"Покупатель с ID {id} удален");

                return true;
            }

            Console.WriteLine($"Покупатель с ID {id} не найден");

            return false;
        }

        public bool DeleteSale(int id)
        {
            var sale = Sales.FirstOrDefault(s => s.Id == id);

            if (sale != null)
            {
                Sales.Remove(sale);
                Console.WriteLine($"Продажа с ID {id} удалена");

                return true;
            }

            Console.WriteLine($"Продажа с ID {id} не найдена");

            return false;
        }

        public void AddAnimal(Animal animal)
        {
            if (Animals.Any(a => a.Id == animal.Id))
            {
                Console.WriteLine("Животное с таким ID уже существует!");

                return;
            }

            Animals.Add(animal);
            Console.WriteLine("Животное добавлено успешно!");
        }

        public void AddCustomer(Customer customer)
        {
            if (Customers.Any(c => c.Id == customer.Id))
            {
                Console.WriteLine("Покупатель с таким ID уже существует!");

                return;
            }

            Customers.Add(customer);
            Console.WriteLine("Покупатель добавлен успешно!");
        }

        public void AddSale(Sale sale)
        {
            if (Sales.Any(s => s.Id == sale.Id))
            {
                Console.WriteLine("Продажа с таким ID уже существует!");

                return;
            }

            if (!Animals.Any(a => a.Id == sale.AnimalId))
            {
                Console.WriteLine("Животное с указанным ID не найдено!");

                return;
            }

            if (!Customers.Any(c => c.Id == sale.CustomerId))
            {
                Console.WriteLine("Покупатель с указанным ID не найден!");

                return;
            }

            Sales.Add(sale);
            Console.WriteLine("Продажа добавлена успешно!");
        }

        public List<Animal> GetSphynxCats()
        {
            return Animals
                .Where(a => a.Species.Equals("Кошка", StringComparison.OrdinalIgnoreCase) &&
                            a.Breed.Equals("Сфинкс", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public decimal GetTotalSalesForCustomersOver30()
        {
            var customerIdsOver30 = Customers
                .Where(c => c.Age > 30)
                .Select(c => c.Id)
                .ToList();

            return Sales
                .Where(s => customerIdsOver30.Contains(s.CustomerId))
                .Sum(s => s.Price);
        }

        public List<string> GetCatSalesWithDetails()
        {
            return (from sale in Sales
                    join animal in Animals on sale.AnimalId equals animal.Id
                    join customer in Customers on sale.CustomerId equals customer.Id
                    where animal.Species.Equals("Кошка", StringComparison.OrdinalIgnoreCase)
                    select $"Дата: {sale.Date.ToShortDateString()}, Порода: {animal.Breed}, " +
                           $"Покупатель: {customer.Name}, Цена: {sale.Price}")
                    .ToList();
        }

        public double GetAverageAgeOfDachshundBuyers()
        {
            var customerIds = (from sale in Sales
                               join animal in Animals on sale.AnimalId equals animal.Id
                               where animal.Species.Equals("Собака", StringComparison.OrdinalIgnoreCase) &&
                                     animal.Breed.Equals("Такса", StringComparison.OrdinalIgnoreCase)
                               select sale.CustomerId)
                              .Distinct()
                              .ToList();

            if (!customerIds.Any())
            {
                return 0;
            }

            return Customers
                .Where(c => customerIds.Contains(c.Id))
                .Average(c => c.Age);
        }

        public void SaveToExcel()
        {
            string tempFile = "";

            try
            {
                // Проверяем, загружен ли файл
                if (string.IsNullOrEmpty(FilePath))
                {
                    Console.WriteLine("Ошибка: не указан путь к файлу!");
                    return;
                }

                // Принудительная сборка мусора для закрытия всех потоков
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // Создаем временный файл
                tempFile = Path.GetTempFileName();

                // Устанавливаем лицензию для некоммерческого использования
                ExcelPackage.License.SetNonCommercialPersonal("ZooShopLINQ");

                using (ExcelPackage package = new ExcelPackage())
                {
                    // === ЛИСТ "Животные" ===
                    ExcelWorksheet wsAnimals = package.Workbook.Worksheets.Add("Животные");

                    wsAnimals.Cells[1, 1].Value = "ID";
                    wsAnimals.Cells[1, 2].Value = "Вид";
                    wsAnimals.Cells[1, 3].Value = "Порода";

                    for (int i = 0; i < Animals.Count; i++)
                    {
                        wsAnimals.Cells[i + 2, 1].Value = Animals[i].Id;
                        wsAnimals.Cells[i + 2, 2].Value = Animals[i].Species;
                        wsAnimals.Cells[i + 2, 3].Value = Animals[i].Breed;
                    }

                    // === ЛИСТ "Покупатели" ===
                    ExcelWorksheet wsCustomers = package.Workbook.Worksheets.Add("Покупатели");

                    wsCustomers.Cells[1, 1].Value = "ID";
                    wsCustomers.Cells[1, 2].Value = "Имя";
                    wsCustomers.Cells[1, 3].Value = "Возраст";
                    wsCustomers.Cells[1, 4].Value = "Адрес";

                    for (int i = 0; i < Customers.Count; i++)
                    {
                        wsCustomers.Cells[i + 2, 1].Value = Customers[i].Id;
                        wsCustomers.Cells[i + 2, 2].Value = Customers[i].Name;
                        wsCustomers.Cells[i + 2, 3].Value = Customers[i].Age;
                        wsCustomers.Cells[i + 2, 4].Value = Customers[i].Address;
                    }

                    // === ЛИСТ "Продажи" ===
                    ExcelWorksheet wsSales = package.Workbook.Worksheets.Add("Продажи");

                    wsSales.Cells[1, 1].Value = "ID";
                    wsSales.Cells[1, 2].Value = "ID животного";
                    wsSales.Cells[1, 3].Value = "ID покупателя";
                    wsSales.Cells[1, 4].Value = "Дата";
                    wsSales.Cells[1, 5].Value = "Цена";

                    for (int i = 0; i < Sales.Count; i++)
                    {
                        wsSales.Cells[i + 2, 1].Value = Sales[i].Id;
                        wsSales.Cells[i + 2, 2].Value = Sales[i].AnimalId;
                        wsSales.Cells[i + 2, 3].Value = Sales[i].CustomerId;

                        // Явно задаем формат даты
                        wsSales.Cells[i + 2, 4].Value = Sales[i].Date;
                        wsSales.Cells[i + 2, 4].Style.Numberformat.Format = "dd.MM.yyyy";

                        wsSales.Cells[i + 2, 5].Value = Sales[i].Price;
                    }

                    // Автоподбор ширины колонок
                    wsAnimals.Cells[wsAnimals.Dimension.Address].AutoFitColumns();
                    wsCustomers.Cells[wsCustomers.Dimension.Address].AutoFitColumns();
                    wsSales.Cells[wsSales.Dimension.Address].AutoFitColumns();

                    // Сохраняем во временный файл
                    FileInfo tempFileInfo = new FileInfo(tempFile);
                    package.SaveAs(tempFileInfo);
                }

                // Если все хорошо, заменяем оригинальный файл временным
                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                }
                File.Move(tempFile, FilePath);

                Console.WriteLine($"✅ Данные сохранены в файл: {FilePath}");
                Console.WriteLine($"   Сохранено: {Animals.Count} животных, {Customers.Count} покупателей, {Sales.Count} продаж");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Ошибка сохранения: {ex.Message}");

                // Если ошибка, удаляем временный файл
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }
}