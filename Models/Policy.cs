using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceNow_XMLGenerator.Models
{
    public class Policy
    {
        public string PolicyNumber { get; set; }

        public List<Driver> Drivers { get; set; }
        public List<Vehicle> Vehicles { get; set; }
    }
}
