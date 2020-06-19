using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class PaymentPlan
    {
        public string PolicyNumber { get; set; }
        public string EftType { get; set; }
        public string Token { get; set; }
        public string AccountNumber { get; set; }
        public string SweepDate { get; set; }
    }
}
