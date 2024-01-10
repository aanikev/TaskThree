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

while (string.IsNullOrEmpty(excelPath))
{
    Console.WriteLine("Повторите ввод пути до файла: ");
    excelPath = Console.ReadLine();
}

while (true)
{
    Console.WriteLine();
    Console.WriteLine("Введите число доступной комманды:\n\n" +
                        "1 - Вывод информации о клиентах заказавших товар\n" +
                        "2 - Изменение контактного лица клиента\n" +
                        "3 - Определить золотого клиента\n\n");
    int command = int.Parse(Console.ReadLine());
    Console.WriteLine();

    switch (command)
    {
        case 1:
            Console.WriteLine("Введите наименование товара:");
            string productName = Console.ReadLine();

            while (string.IsNullOrEmpty(productName))
            {
                Console.WriteLine("Повторите ввод наименованиz товара: ");
                excelPath = Console.ReadLine();
            }
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
            Console.WriteLine("Введите дату в формате: месяц и год - мм.гггг или год - гггг");
            var readDate = Console.ReadLine();
            Console.WriteLine();

            while (string.IsNullOrEmpty(readDate))
            {
                Console.WriteLine("Повторите ввод даты: ");
                excelPath = Console.ReadLine();
            }
            Console.WriteLine();
            var goldCustomer = new GoldCustomer(readDate, excelPath);
            Console.WriteLine(goldCustomer.GetGoldCustomer());
            break;
        default:
            Console.WriteLine("Вы не выбрани ни одной комманды");
            break;
    }
}

