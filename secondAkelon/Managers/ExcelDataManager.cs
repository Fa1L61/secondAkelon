using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using secondAkelon.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace secondAkelon.Managers
{
    public class ExcelDataManager
    {
        public List<Order> Orders = new List<Order>();
        public List<Customer> Customers = new List<Customer>();
        public List<Product> Products = new List<Product>();

        public ExcelDataManager(string excelPath) 
        {
            using (var workbook = new XLWorkbook(excelPath))
            {
                //Заявки
                var worksheet = workbook.Worksheet(3);
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    Orders.Add(new Order
                    {
                        Id = row.Cell(1).GetValue<int>(),
                        ProductId = row.Cell(2).GetValue<int>(),
                        CustomerId = row.Cell(3).GetValue<int>(),
                        NumberOrder = row.Cell(4).GetValue<int>(),
                        Count = row.Cell(5).GetValue<int>(),
                        OrderDate = DateOnly.Parse(row.Cell(6).GetDateTime().ToString("d"))
                    });
                }
                //Клиенты
                worksheet = workbook.Worksheet(2);
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    Customers.Add(new Customer
                    {
                        Id = row.Cell(1).GetValue<int>(),
                        Name = row.Cell(2).GetValue<string>(),
                        Address = row.Cell(3).GetValue<string>(),
                        ContactName = row.Cell(4).GetValue<string>(),
                    });
                }
                //Продукты
                worksheet = workbook.Worksheet(1);
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    Products.Add(new Product
                    {
                        Id = row.Cell(1).GetValue<int>(),
                        Name = row.Cell(2).GetValue<string>().ToLower(),
                        Unit = row.Cell(3).GetValue<string>(),
                        Cost = row.Cell(4).GetValue<float>(),
                    });
                }
            }
        }
        public void CustomerInfo(string productName)
        {
            var productInfo = Products.Where(i => i.Name == productName.ToLower()).First();
            var ordersInfo = Orders.Where(i => i.ProductId == productInfo.Id).ToList();

            var customers = new List<Customer>();
            foreach (var order in ordersInfo)
            {
                customers.Add(Customers.Where(i => i.Id == order.CustomerId).First());
            }

            Console.WriteLine($"Клиенты, заказавшие {productName.ToLower()}:\n");
            for (int i = 0; i < customers.Count; i++)
            {
                Console.WriteLine($"{customers[i].Name}.\nДоставка по адресу: {customers[i].Address}.\nКонтактное лицо: {customers[i].ContactName}");
                Console.WriteLine($"Требуемое количество: {ordersInfo[i].Count} {productInfo.Unit}");
                Console.WriteLine($"Стоимость за единицу: {productInfo.Cost} Руб.");
                Console.WriteLine($"Общая сумма заказа: {productInfo.Cost * ordersInfo[i].Count} Руб.");
                Console.WriteLine($"Дата заказа: {ordersInfo[i].OrderDate}");

                Console.WriteLine();
            }
        }

        public void UpdateCustomer(string productName)
        {

        }
    }
}
