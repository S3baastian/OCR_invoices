using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OCR
{
    public class Row
    {
        public Row() { }
        public Row(string code, string description, double quantity, double price)
        {
            Code = code;
            Description = description;
            Quantity = quantity;
            Price = price;
        }
        public string Code { get; set; }
        public string Description { get; set; }
        public double Price { get; set; }
        public double Quantity { get; set; }
        public string Discount { get; set; }
        public double TotalValue { get; set; }

        public void DisplayRows()
        {
            Console.Write(Code + "\t" + Description + "\t" + Quantity + "\t" + Price + "\t" + Discount + "\t" + TotalValue +"\n");


        }
    }
}
