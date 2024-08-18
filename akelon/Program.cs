using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public class Product
{
    public string ProductCode { get; set; }
    public string ProductName { get; set; }
    public string UnitOfMeasure { get; set; }
    public float Price { get; set; }
}

public class Client
{
    public string ClientCode { get; set; }
    public string OrganizationName { get; set; }
    public string Address { get; set; }
    public string ContactPerson { get; set; }
}

public class Order
{
    public string OrderCode { get; set; }
    public string ProductCode { get; set; }
    public string ClientCode { get; set; }
    public string OrderNumber { get; set; }
    public int Quantity { get; set; }
    public DateTime OrderDate { get; set; }
}

class Program
{
    public static void Main(string[] args)
    {
        Console.WriteLine("Введите путь до файла с данными:");
        string filePath = Console.ReadLine();

        if (!System.IO.File.Exists(filePath))
        {
            Console.WriteLine("Файл не существует по указанному пути.");
            return;
        }

        try
        {
            // Считываем все листы в документе.
            var products = new Dictionary<string, Product>();
            var clients = new Dictionary<string, Client>();
            var orders = new Dictionary<string, Order>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                products = ReadProducts(document);
                clients = ReadClients(document);
                orders = ReadOrders(document);
            }

            Console.WriteLine("Выберите команду:");
            Console.WriteLine("1 - По наименованию товара вывести информацию о клиентах");
            Console.WriteLine("2 - Изменить контактное лицо клиента");
            Console.WriteLine("3 - Определить золотого клиента за указанный год и месяц");

            int command = int.Parse(Console.ReadLine());

            switch (command)
            {
                case 1:
                    GetProductsOrdersByClients(products, clients, orders);
                    break;
                case 2:
                    ChangeContact(clients, filePath);
                    break;
                case 3:
                    GetGoldenClient(orders);
                    break;
                default:
                    Console.WriteLine("Неизвестная команда");
                    break;
            }
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
        }
    }

    public static SheetData ReadSheetElementsFromList(SpreadsheetDocument document, string listName)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => GetSheetName(workbookPart, s.Id) == listName);

        if (sheet == null)
            throw new Exception("Лист с именем 'Товары' не найден");

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        return worksheetPart.Worksheet.GetFirstChild<SheetData>();
    }

    public static Dictionary<string, Product> ReadProducts(SpreadsheetDocument document)
    {
        var products = new Dictionary<string, Product>();

        var sheetData = ReadSheetElementsFromList(document, "Товары");

        foreach (var row in sheetData.Elements<Row>().Skip(1)) // Пропускаем заголовок
        {
            var cells = row.Elements<Cell>().Select(c => GetCellValue(document, c)).ToList();
            if (cells.Count >= 4 && !string.IsNullOrEmpty(cells[0]) && !string.IsNullOrEmpty(cells[3]))
            {
                var product = new Product
                {
                    ProductCode = cells[0],
                    ProductName = cells[1],
                    UnitOfMeasure = cells[2],
                    Price = float.Parse(cells[3].Replace(" ₽", "").Replace(",", "."), CultureInfo.InvariantCulture)
                };

                products[product.ProductCode] = product;
            }
        }

        return products;
    }

    public static Dictionary<string, Client> ReadClients(SpreadsheetDocument document)
    {
        var clients = new Dictionary<string, Client>();

        var sheetData = ReadSheetElementsFromList(document, "Клиенты");

        foreach (var row in sheetData.Elements<Row>().Skip(1)) // Пропускаем заголовок
        {
            var cells = row.Elements<Cell>().Select(c => GetCellValue(document, c)).ToList();
            if (cells.Count >= 4 && !string.IsNullOrEmpty(cells[0]))
            {
                var client = new Client
                {
                    ClientCode = cells[0],
                    OrganizationName = cells[1],
                    Address = cells[2],
                    ContactPerson = cells[3]
                };

                clients[client.ClientCode] = client;
            }
        }

        return clients;
    }

    public static Dictionary<string, Order> ReadOrders(SpreadsheetDocument document)
    {
        var orders = new Dictionary<string, Order>();

        var sheetData = ReadSheetElementsFromList(document, "Заявки");

        foreach (var row in sheetData.Elements<Row>().Skip(1)) // Пропускаем заголовок
        {
            var cells = row.Elements<Cell>().Select(c => GetCellValue(document, c)).ToList();
            var baseDate = new DateTime(1900, 1, 1);
            var isDateParceCorrect = double.TryParse(cells[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double orderDateDouble);
            if (cells.Count >= 6 && !string.IsNullOrEmpty(cells[0]) && !string.IsNullOrEmpty(cells[1]) &&
                !string.IsNullOrEmpty(cells[2]) && int.TryParse(cells[4], out int quantity) && isDateParceCorrect)
            {
                var order = new Order
                {
                    OrderCode = cells[0],
                    ProductCode = cells[1],
                    ClientCode = cells[2],
                    OrderNumber = cells[3],
                    Quantity = quantity,
                    OrderDate = baseDate.AddDays(orderDateDouble - 2)
                };

                orders[order.OrderCode] = order;
            }
        }

        return orders;
    }

    public static string GetSheetName(WorkbookPart workbookPart, string sheetId)
    {
        var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Id.Value == sheetId);
        return sheet?.Name?.Value ?? "Unknown sheet name";
    }

    public static string GetCellValue(SpreadsheetDocument document, Cell cell)
    {
        if (cell == null || cell.CellValue == null)
            return null;

        var value = cell.CellValue.InnerText;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var stringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTablePart != null)
            {
                var sharedString = stringTablePart.SharedStringTable.ElementAtOrDefault(int.Parse(value));
                value = sharedString?.InnerText ?? string.Empty;
            }
        }
        return value;
    }

    public static void GetProductsOrdersByClients(Dictionary<string, Product> products, Dictionary<string, Client> clients, Dictionary<string, Order> orders)
    {
        Console.WriteLine("Введите наименование товара:");
        string productName = Console.ReadLine();
        var product = products.Values.FirstOrDefault(p => p.ProductName.Equals(productName, StringComparison.OrdinalIgnoreCase));

        if (product == null)
        {
            Console.WriteLine("Товар не найден.");
            return;
        }

        Console.WriteLine($"Информация о клиентах, заказавших товар {productName}:");
        var orderDetails = orders.Values.Where(o => o.ProductCode == product.ProductCode).ToList();

        foreach (var order in orderDetails)
        {
            var client = clients.GetValueOrDefault(order.ClientCode);
            if (client != null)
            {
                Console.WriteLine($"Клиент: {client.OrganizationName}, Количество: {order.Quantity}, Цена: {product.Price} ₽, Дата: {order.OrderDate.ToShortDateString()}");
            }
        }
    }

    public static void ChangeContact(Dictionary<string, Client> clients, string filePath)
    {
        Console.WriteLine("Введите название организации:");
        string organizationName = Console.ReadLine();
        Console.WriteLine("Введите ФИО нового контактного лица:");
        string newContactName = Console.ReadLine();

        var client = clients.Values.FirstOrDefault(c => c.OrganizationName.Equals(organizationName, StringComparison.OrdinalIgnoreCase));

        if (client == null)
        {
            Console.WriteLine("Организация не найдена.");
            return;
        }

        client.ContactPerson = newContactName;

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
        {
            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => GetSheetName(workbookPart, s.Id) == "Клиенты");

            if (sheet == null)
                throw new Exception("Лист с именем 'Клиенты' не найден");

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var rows = sheetData.Elements<Row>().Skip(1).ToList();

            foreach (var row in rows)
            {
                var cells = row.Elements<Cell>().Select(c => GetCellValue(document, c)).ToList();
                if (cells.Count >= 4 && cells[1].Equals(organizationName, StringComparison.OrdinalIgnoreCase))
                {
                    var contactCell = row.Elements<Cell>().ElementAtOrDefault(3);
                    if (contactCell != null)
                    {
                        contactCell.CellValue = new CellValue(newContactName);
                        contactCell.DataType = CellValues.String;
                        worksheetPart.Worksheet.Save();
                    }
                    break;
                }
            }

            Console.WriteLine("Контактное лицо успешно обновлено.");
        }
    }

    public static void GetGoldenClient(Dictionary<string, Order> orders)
    {
        Console.WriteLine("Введите год (целиком):");
        int year = int.Parse(Console.ReadLine());
        Console.WriteLine("Введите месяц (числом):");
        int month = int.Parse(Console.ReadLine());

        var monthlyOrders = orders.Values.Where(o => o.OrderDate.Year == year && o.OrderDate.Month == month).ToList();
        var clientOrderCounts = monthlyOrders.GroupBy(o => o.ClientCode)
            .Select(g => new
            {
                ClientCode = g.Key,
                OrderCount = g.Count()
            })
            .OrderByDescending(c => c.OrderCount)
            .FirstOrDefault();

        if (clientOrderCounts != null)
        {
            Console.WriteLine($"Золотой клиент: Код клиента {clientOrderCounts.ClientCode}, Количество заказов: {clientOrderCounts.OrderCount}");
        }
        else
        {
            Console.WriteLine("Нет данных для указанных периода.");
        }
    }
}