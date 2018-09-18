using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataSources
{
    public class Customer
    {
        public string Name { get; set; }
        public string RegisterDate { get; set; }
        public string LastBuy { get; set; }
        public string Item { get; set; }
        public int Quantity { get; set; }
        public double ItemCost { get; set; }
    }
}
