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
    }
}
