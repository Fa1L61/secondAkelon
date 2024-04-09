using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace secondAkelon.Models
{
    public class Order
    {
        public int Id { get; set; }
        public int ProductId {  get; set; }
        public int CustomerId { get; set; }
        public int NumberOrder { get; set; }
        public int Count { get; set; }
        public DateOnly OrderDate { get; set; }

        //public Order(int id, int productId, int customerId, int numberOrder, int count, DateOnly orderDate) 
        //{ 
        //    Id = id;
        //    ProductId = productId;
        //    CustomerId = customerId;
        //    NumberOrder = numberOrder;
        //    Count = count;
        //    OrderDate = orderDate;

        //}
    }
}
