using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


static class DataManager
{
    static List<Client> clients = new();
    static List<Order> orders = new();
    static List<Service> services = new();
    static List<ServiceType> serviceTypes = new();

    // Вписываем имя файла и на выходе получаем 4 коллекции со всеми данными таблицы
    public static void LoadFromExcel()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage(new FileInfo("LR5-var7.xlsx"));

        clients.Clear();
        orders.Clear();
        services.Clear();
        serviceTypes.Clear();

        var wsClients = package.Workbook.Worksheets["Клиенты"];
        var wsOrders = package.Workbook.Worksheets["Заказы"];
        var wsServices = package.Workbook.Worksheets["Услуги"];
        var wsTypes = package.Workbook.Worksheets["Типы услуг"];

        for (int i = 2; i <= wsClients.Dimension.End.Row; i++)
            clients.Add(new Client(
                    int.Parse(wsClients.Cells[i, 1].Text), 
                    wsClients.Cells[i, 2].Text,            
                    wsClients.Cells[i, 3].Text,            
                    wsClients.Cells[i, 4].Text,            
                    wsClients.Cells[i, 5].Text             
            ));
        

        for (int i = 2; i <= wsOrders.Dimension.End.Row; i++)
        {
            var cellValue = wsOrders.Cells[i, 3].Value;

            DateTime date;

            if (cellValue is DateTime dt)
                date = dt;
            else
                date = DateTime.FromOADate(Convert.ToDouble(cellValue)); 


            orders.Add(new Order(
                int.Parse(wsOrders.Cells[i, 1].Text),
                int.Parse(wsOrders.Cells[i, 2].Text),
                int.Parse(wsOrders.Cells[i, 4].Text), 
                date,
                int.Parse(wsOrders.Cells[i, 5].Text) 
            ));
        }

        for (int i = 2; i <= wsServices.Dimension.End.Row; i++)
            services.Add(new Service(
                int.Parse(wsServices.Cells[i, 1].Text),
                int.Parse(wsServices.Cells[i, 2].Text),
                wsServices.Cells[i, 3].Text,
                wsServices.Cells[i, 4].Text
            ));

        for (int i = 2; i <= wsTypes.Dimension.End.Row; i++)
            serviceTypes.Add(new ServiceType(
                int.Parse(wsTypes.Cells[i, 1].Text),
                wsTypes.Cells[i, 2].Text));
    }

    // Получаем название листа и перебираем его в коллекции
    public static void ShowSheet(string sheetName)
    {
        switch (sheetName.ToLower())
        {
            case "клиенты":
                foreach (var c in clients) Console.WriteLine(c);
                break;
            case "заказы":
                foreach (var o in orders) Console.WriteLine(o);
                break;
            case "услуги":
                foreach (var s in services) Console.WriteLine(s);
                break;
            case "типы услуг":
                foreach (var t in serviceTypes) Console.WriteLine(t);
                break;
            default:
                Console.WriteLine("Такого листа нет.");
                break;
        }
    }

    // Получаем название листа и его айди, находим в коллекции и удаляем
    public static void RemoveById(string sheetName, int id)
    {
        bool removed = sheetName.ToLower() switch
        {
            "клиенты" => clients.RemoveAll(c => c.Id == id) > 0,
            "заказы" => orders.RemoveAll(o => o.Id == id) > 0,
            "услуги" => services.RemoveAll(s => s.Id == id) > 0,
            "типы услуг" => serviceTypes.RemoveAll(t => t.Id == id) > 0,
            _ => false
        };

        Console.WriteLine(removed ? "Элемент удалён." : "Элемент с таким Id не найден.");
    }

    // Получаем название листа и список со всеми полями в правильном порядке
    public static void AddToSheet(string sheetName, params string[] values)
    {
        try
        {
            switch (sheetName.ToLower())
            {
                case "клиенты":
                    clients.Add(new Client(
                        int.Parse(values[0]),
                        values[1],
                        values[2],
                        values[3],
                        values[4]
                    ));
                    break;

                case "заказы":
                    orders.Add(new Order(
                        int.Parse(values[0]),
                        int.Parse(values[1]),
                        int.Parse(values[2]),
                        DateTime.Parse(values[3]),
                        int.Parse(values[4])
                    ));
                    break;

                case "услуги":
                    services.Add(new Service(
                        int.Parse(values[0]),
                        int.Parse(values[1]),
                        values[2],
                        values[3]
                    ));
                    break;

                case "типы услуг":
                    serviceTypes.Add(new ServiceType(
                        int.Parse(values[0]),
                        values[1]
                    ));
                    break;

                default:
                    Console.WriteLine("Такого листа нет.");
                    return;
            }

            Console.WriteLine("Элемент добавлен в коллекцию. Не забудьте сохранить изменения в Excel.");
        }
        catch
        {
            Console.WriteLine("Ошибка при добавлении элемента. Проверьте данные.");
        }
    }

    // LINQ Запросы

    // Перебираем через Where и ищем совпадения
    public static void QueryClientsFromCity(string city)
    {
        var result = clients.Where(c => c.City == city);
        foreach (var c in result) Console.WriteLine(c);
    }
      
    // Перебираем через Count и получаем ответ
    public static void QueryOrdersCountByClient()
    {
        Console.Write("Введите код клиента: ");
        int id = int.Parse(Console.ReadLine());
        int count = orders.Count(o => o.ClientId == id);
        Console.WriteLine("Количество заказов: " + count);
    }

    // Соединяем все таблицы и выводим нужные сообщения
    public static void QueryOrdersWithServices()
    {
        var result = from o in orders
                        join c in clients on o.ClientId equals c.Id
                        join s in services on o.ServiceId equals s.Id
                        select new
                        {
                            OrderId = o.Id,
                            Client = c.LastName,
                            Service = s.Name,
                            Total = s.Price * o.Quantity
                        };

        foreach (var r in result) 
            Console.WriteLine($"Заказ {r.OrderId} | {r.Client} | {r.Service} | {r.Total}");
    }

    // Соединяем таблицы, ищем совпадения и считаем все цены в список, и суммируем их
    public static void QueryPolygraphySum()
    {
        // фильтруем только заказы из Владивостока
        var filteredOrders = from o in orders
                             join c in clients on o.ClientId equals c.Id
                             join s in services on o.ServiceId equals s.Id
                             join t in serviceTypes on s.TypeId equals t.Id
                             where c.City == "Владивосток"
                                   && t.Name== "Полиграфия"
                                   && o.Date >= new DateTime(2018, 6, 1)
                                   && o.Date <= new DateTime(2018, 6, 30)
                             select (decimal)s.Price * o.Quantity;

        decimal sum = filteredOrders.Sum();

        Console.WriteLine("Общая стоимость: " + sum);

    }

    // Удаляем старые значения в листе и вписываем из коллекции
    public static void SaveToExcel()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage(new FileInfo("LR5-var7.xlsx"));

        SaveClients(package.Workbook.Worksheets["Клиенты"]);
        SaveOrders(package.Workbook.Worksheets["Заказы"]);
        SaveServices(package.Workbook.Worksheets["Услуги"]);
        SaveServiceTypes(package.Workbook.Worksheets["Типы услуг"]);

        package.Save();
    }

    static void SaveClients(ExcelWorksheet ws)
    {
        ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Clear();

        int row = 2;
        foreach (var c in clients)
        {
            ws.Cells[row, 1].Value = c.Id;
            ws.Cells[row, 2].Value = c.LastName;
            ws.Cells[row, 3].Value = c.FirstName;
            ws.Cells[row, 4].Value = c.Patronymic;
            ws.Cells[row, 5].Value = c.City.StartsWith("г. ") ? c.City : "г. " + c.City;
            row++;
        }
    }

    static void SaveOrders(ExcelWorksheet ws)
    {
        ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Clear();

        int row = 2;
        foreach (var o in orders)
        {
            ws.Cells[row, 1].Value = o.Id;
            ws.Cells[row, 2].Value = o.ClientId;
            ws.Cells[row, 3].Value = o.Date;
            ws.Cells[row, 3].Style.Numberformat.Format = "dd.MM.yyyy";
            ws.Cells[row, 4].Value = o.ServiceId;
            ws.Cells[row, 5].Value = o.Quantity;
            row++;
        }
    }

    static void SaveServices(ExcelWorksheet ws)
    {
        ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Clear();

        int row = 2;
        foreach (var s in services)
        {
            ws.Cells[row, 1].Value = s.Id;
            ws.Cells[row, 2].Value = s.TypeId;
            ws.Cells[row, 3].Value = s.Name;
            ws.Cells[row, 4].Value = s.Price + " р.";
            row++;
        }
    }

    static void SaveServiceTypes(ExcelWorksheet ws)
    {
        ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Clear();

        int row = 2;
        foreach (var t in serviceTypes)
        {
            ws.Cells[row, 1].Value = t.Id;
            ws.Cells[row, 2].Value = t.Name;
            row++;
        }
    }

}

// Классы для коллекций
class Client
{
    public int Id { get; set; }
    public string LastName { get; set; }
    public string FirstName { get; set; }
    public string Patronymic { get; set; }
    public string City { get; set; }

    public Client(int id, string lastName, string firstName, string patronymic, string city)
    {
        Id = id;
        LastName = lastName;
        FirstName = firstName;
        Patronymic = patronymic;
        City = city.StartsWith("г. ") ? city.Substring(3) : city;
    }

    public override string ToString()
    {
        return $"Клиент {Id}: {LastName} {FirstName} {Patronymic}, город {City}";
    }
}


class Order
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public int ServiceId { get; set; }
    public DateTime Date { get; set; }
    public int Quantity { get; set; }

    public Order(int id, int clientId, int serviceId, DateTime date, int quantity)
    {
        Id = id;
        ClientId = clientId;
        ServiceId = serviceId;
        Date = date;
        Quantity = quantity;
    }

    public override string ToString()
    {
        return $"Заказ {Id}: клиент {ClientId}, услуга {ServiceId}, дата {Date:d}, кол-во {Quantity}";
    }
}

class Service
{
    public int Id { get; set; }
    public int TypeId { get; set; }
    public string Name { get; set; }
    public decimal Price { get; set; }

    public Service(int id, int typeId, string name, string priceString)
    {
        Id = id;
        TypeId = typeId;
        Name = name;
        if (!string.IsNullOrEmpty(priceString))
        {
            var cleaned = priceString.Replace(" р.", "").Trim();
            Price = decimal.Parse(cleaned);
        }
    }

    public override string ToString()
    {
        return $"Услуга {Id}: {Name}, цена {Price} р.";
    }
}

class ServiceType
{
    public int Id { get; set; }
    public string Name { get; set; }

    public ServiceType(int id, string name)
    {
        Id = id;
        Name = name;
    }

    public override string ToString()
    {
        return $"Тип услуги {Id}: {Name}";
    }
}