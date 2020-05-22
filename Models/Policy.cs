using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceNow_XMLGenerator.Models
{
    public class Policy
    {
        public General General { get; set; }
        public string PolicyNumber { get; set; }
        public List<Driver> Drivers { get; set; }
        public List<Vehicle> Vehicles { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string BillReminder { get; set; }
        public string BirthDate { get; set; }
        public string PrimaryPhone { get; set; }
        public string MailingAddress { get; set; }
        public string MailingCity { get; set; }
        public string MailingState { get; set; }
        public string MailingZip { get; set; }
        public string GaragingAddress { get; set; }
        public string GaragingCity { get; set; }
        public string GaragingState { get; set; }
        public string GaragingZip { get; set; }
        public string InceptionDate { get; set; }
        public string EffectiveDate { get; set; }
        public string ExpirationDate { get; set; }
        public string Term { get; set; }
        public string ProducerCode { get; set; }
        public string Age { get; set; }
    }
}
