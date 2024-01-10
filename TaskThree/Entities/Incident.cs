using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace TaskThree.Entities
{
    internal class Incident
    {
        public int IncidentCode { get; set; }
        public int ProductCode { get; set; }
        public int CustomerCode { get; set; }
        public int NumberIncident { get; set; }
        public int RequiredQuantity { get; set; }
        public DateTime PostingDate { get; set; }
    }
}
