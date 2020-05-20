using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceNow_XMLGenerator.Models
{
    public class Vehicle
    {
        public string PolicyNumber { get; set; }
        public string VIN { get; set; }
        public string CollDeduct { get; set; }
        public string CompDeduct { get; set; }
        public string AnnualMileage { get; set; }
        public string CurrentOdometer { get; set; }
        public string Lessor { get; set; }
        public string Rental { get; set; }
        public string No { get; set; }
        public string Pub { get; set; }
        public string Dr { get; set; }
    }
}
