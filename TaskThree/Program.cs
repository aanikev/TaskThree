using System.Collections.Immutable;
using System.Globalization;
using System.IO;
using System.Reflection.Metadata.Ecma335;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using TaskThree;
using TaskThree.Entities;
using static ClosedXML.Excel.XLPredefinedFormat;

Console.WriteLine("Введите путь до файла: ");
string excelPath = Console.ReadLine();
Console.WriteLine();
Console.WriteLine("Введите число доступной комманды:");
Console.WriteLine();
Console.WriteLine("1 - Вывод информации о клиентах заказавших товар");
Console.WriteLine("2 - Изменение контактного лица клиента");
Console.WriteLine("3 - Определить золотого клиента");
Console.WriteLine();
int command = int.Parse(Console.ReadLine());
Console.WriteLine();
switch (command)
{
    case 1:
        Console.WriteLine("Введите наименование товара:");
        string productName = Console.ReadLine();
        Console.WriteLine();
        var customerInfo = new CustomerInfo(productName, excelPath);
        var rezult = customerInfo.GetCustomerInfo();
        Console.WriteLine(rezult);
        break;
    case 2:
        Console.WriteLine("Укажите нименование организации изменяемого контактного лица:");
        string organizationName = Console.ReadLine();
        Console.WriteLine("Укажите ФИО изменяемого контактного лица:");
        string newContactPerson = Console.ReadLine();
        var editContactPerson = new EditContactPerson(organizationName, newContactPerson, excelPath);
        var reault = editContactPerson.EditContactInTable();
        Console.WriteLine(reault);
        break;
    case 3:
        Console.WriteLine("Введите дату в формате: месяц и год - XX.XXXX или год - ХХХХ");
        var readDate = Console.ReadLine();
        Console.WriteLine();
        var goldCustomer = new GoldCustomer(readDate, excelPath);
        Console.WriteLine(goldCustomer.GetGoldCustomer());
        break;
    default:
        Console.WriteLine("Вы не выбрани ни одной комманды");
        break;
}

