using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskThree.Entities
{
    internal class Customer
    {
        public int CustomerCode { get; set; }
        public string OrganizationName { get; set; }
        public string OrganizationAddress { get; set; }
        public string ContactPerson { get; set; }
    }
}
