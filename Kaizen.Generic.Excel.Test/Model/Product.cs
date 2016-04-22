using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kaizen.Generic.Excel.Test.Model
{
    public class Product
    {
        public Product(string name, double price)
        {
            Name = name;
            Price = price;
        }

        [DisplayName("Nome")]
        public string Name { get; set; }
        [DisplayName("Preço")]
        public double Price { get; set; }

    }
}
