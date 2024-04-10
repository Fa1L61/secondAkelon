using System;
using System.IO;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using secondAkelon.Managers;
using secondAkelon.Models;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelDataManager excelDataManager = null;

            string path = "";

            while (true)
            {
                Console.WriteLine("Добро пожаловать в главное меню\nДля навигации используйте:");
                Console.WriteLine("1 - для задания ввода пути к Excel файлу");
                Console.WriteLine("2 - для вывода информации о клиентах, купивших определенный товар");
                Console.WriteLine("3 - для изменения контактного лица компании");
                Console.WriteLine("4 - для вывода золотого клиента в определенный месяц/год");
                Console.WriteLine("q - для выхода");
                Console.WriteLine();

                var input = Console.ReadLine();

                switch (input)
                {
                    case "1":

                        Console.WriteLine("Введите путь к Вашему файлу: ");
                        path = Console.ReadLine();

                        try
                        {
                            excelDataManager = new ExcelDataManager(path);
                            Console.WriteLine("\nФайл успешно открыт!\n");

                        }
                        catch
                        {
                            Console.WriteLine("\nВы указали некорректный путь до файла. Возвращаюсь в главное меню\n");

                        }
                        break; 
                    case "2":
                        if (excelDataManager == null)
                        {
                            Console.WriteLine("\nВы не подключили Excel файл к приложению. Возвращаюсь в главное меню\n");

                        }
                        else
                        {
                            Console.WriteLine("\nВведите продукт, по которому хотите получить информацию о заказах:");
                            excelDataManager.CustomerInfo(Console.ReadLine());
                        }
                        break; 
                    case "3":
                        if (excelDataManager == null)
                        {
                            Console.WriteLine("\nВы не подключили Excel файл к приложению. Возвращаюсь в главное меню\n");
                        }
                        else
                        {
                            Console.WriteLine("Введите строку в формате \"{Название компании}\" {Фамилия Имя Отчество} нового контактного лица\nНапример:");
                            Console.WriteLine("\"ООО Снег\" Иванов Иван Иванович");
                            excelDataManager.UpdateCustomerContact(Console.ReadLine(), path);
                        }
                        break; 
                    case "4":
                        if (excelDataManager == null)
                        {
                            Console.WriteLine("\nВы не подключили Excel файл к приложению. Возвращаюсь в главное меню\n");
                        }
                        else
                        {
                            Console.WriteLine("\nВведите:");
                            Console.WriteLine("- через точку месяц и год, в период которого хотите найти золотого клиента для поиска по году и месяцу (Пример: 3.2023");
                            Console.WriteLine("- год, за который хотите найти золотого клиента (Пример: 2023");
                            
                            string date = Console.ReadLine();

                            excelDataManager.TopCustomer(date);
                        }
                        break;
                    case "q":
                        Console.WriteLine("\nВсего доброго!\n");
                        Thread.Sleep(1000);
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("\nВы ввели некорректный номер...\n");
                        break;

                }
            }
        }
    }
}