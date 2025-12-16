using System;

class Program
{
    static void Main()
    {
        try
        {
            DataManager.LoadFromExcel();
            Console.WriteLine("База данных загружена.\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ошибка загрузки Excel: " + ex.Message);
            return;
        }

        while (true)
        {
            Console.WriteLine("""
                --- Меню ---
                1 – Просмотр листа
                2 – Добавить элемент в лист
                3 – Удалить элемент по Id
                4 – LINQ-запрос: клиенты из города
                5 – LINQ-запрос: количество заказов клиента
                6 – LINQ-запрос: заказы + услуги
                7 – LINQ-запрос: стоимость Полиграфии
                8 – Сохранить изменения в Excel
                0 – Выход
                """);

            Console.Write("Выбор: ");
            string choice = Console.ReadLine()?.Trim();

            switch (choice)
            {
                case "1":
                    Console.Write("Введите название листа (Клиенты, Заказы, Услуги, Типы услуг): ");
                    string sheetView = Console.ReadLine()?.Trim();
                    DataManager.ShowSheet(sheetView);
                    break;

                case "2":
                    Console.Write("Введите название листа (Клиенты, Заказы, Услуги, Типы услуг): ");
                    string sheetAdd = Console.ReadLine()?.Trim();
                    Console.WriteLine("Введите значения через запятую (все колонки):");
                    string inputAdd = Console.ReadLine() ?? "";
                    string[] valuesAdd = inputAdd.Split(',', StringSplitOptions.TrimEntries);
                    DataManager.AddToSheet(sheetAdd, valuesAdd);
                    break;

                case "3":
                    Console.Write("Введите название листа (Клиенты, Заказы, Услуги, Типы услуг): ");
                    string sheetDel = Console.ReadLine()?.Trim();
                    Console.Write("Введите Id элемента для удаления: ");
                    if (int.TryParse(Console.ReadLine(), out int delId))
                        DataManager.RemoveById(sheetDel, delId);
                    else
                        Console.WriteLine("Неверный Id.");
                    break;

                case "4":
                    Console.Write("Город для поиска клиентов: ");
                    string city = Console.ReadLine()?.Trim();
                    DataManager.QueryClientsFromCity(city);
                    break;

                case "5":
                    DataManager.QueryOrdersCountByClient();
                    break;

                case "6":
                    DataManager.QueryOrdersWithServices();
                    break;

                case "7":
                    DataManager.QueryPolygraphySum();
                    break;

                case "8":
                    DataManager.SaveToExcel();
                    Console.WriteLine("Изменения сохранены в Excel.");
                    break;

                case "0":
                    Console.WriteLine("Выход...");
                    return;

                default:
                    Console.WriteLine("Неверный пункт меню, попробуйте снова.");
                    break;
            }

            Console.WriteLine();
        }
    }
}
