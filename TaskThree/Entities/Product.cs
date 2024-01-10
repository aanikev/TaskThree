using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskThree.Entities
{
    internal class Product
    {
        public int ProductCode { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public double ProductPricePerUnit { get; set; }
    }
}
