using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceNow_XMLGenerator.Models
{
    public class VehicleCoverages
    {
        public Coverage COLL { get; set; }
        public Coverage CPR { get; set; }
        public Coverage Rental { get; set; }

        public VehicleCoverages(string _COLL, string _CPR, string _Rental)
        {
            COLL = new Coverage() { InputValue = _COLL, hasCoverage = false };
            CPR = new Coverage() { InputValue = _CPR, hasCoverage = false };
            Rental = new Coverage() { InputValue = _Rental, hasCoverage = false };
        }

        public void ProcessCoverages()
        {
            ProcessCOLL();
            ProcessCPR();
            ProcessRental();
        }

        private void ProcessCOLL()
        {
            if (!string.IsNullOrEmpty(COLL.InputValue))
            {
                COLL.hasCoverage = true;
                COLL.Value1 = COLL.InputValue;
            }
        }

        private void ProcessCPR()
        {
            if (!string.IsNullOrEmpty(CPR.InputValue))
            {
                CPR.hasCoverage = true;
                CPR.Value1 = CPR.InputValue;
            }
        }

        private void ProcessRental()
        {
            if (!string.IsNullOrEmpty(Rental.InputValue) && Rental.InputValue.Split('/').Length == 2)
            {
                Rental.hasCoverage = true;
                string[] limits = Rental.InputValue.Split('/');
                limits[0] = limits[0] + "000";
                limits[1] = limits[1] + "000";
                string newInputValue = string.Format("{0}/{1}", limits[0], limits[1]);

                Rental.InputValue = newInputValue;
                Rental.Value1 = limits[0];
                Rental.Value2 = limits[1];
            }
        }
    }
}
