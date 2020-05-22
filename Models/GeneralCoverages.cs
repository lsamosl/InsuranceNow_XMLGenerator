using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceNow_XMLGenerator.Models
{
    public class GeneralCoverages
    {
        public Coverage BI { get; set; }
        public Coverage PD { get; set; }
        public Coverage UMBI { get; set; }
        public Coverage MED { get; set; }

        public GeneralCoverages(string _BI, string  _PD, string _UMBI, string _MED)
        {
            BI = new Coverage() { InputValue = _BI, hasCoverage = false };
            PD = new Coverage() { InputValue = _PD, hasCoverage = false };
            UMBI = new Coverage() { InputValue = _UMBI, hasCoverage = false };
            MED = new Coverage() { InputValue = _UMBI, hasCoverage = false };
        }

        public void ProcessCoverages()
        {
            ProcessBI();
            ProcessPD();
            ProcessUMBI();
            ProcessMed();
        }

        private void ProcessBI()
        {
            if (!string.IsNullOrEmpty(BI.InputValue) && BI.InputValue.Split('/').Length == 2)
            {
                BI.hasCoverage = true;
                string[] limits = BI.InputValue.Split('/');
                limits[0] = limits[0] + "000";
                limits[1] = limits[1] + "000";
                string newInputValue = string.Format("{0}/{1}", limits[0], limits[1]);

                BI.InputValue = newInputValue;
                BI.Value1 = limits[0];
                BI.Value2 = limits[1];
            }
        }

        private void ProcessPD()
        {
            if (!string.IsNullOrEmpty(PD.InputValue))
            {
                PD.hasCoverage = true;
                PD.Value1 = PD.InputValue;
            }
        }

        private void ProcessUMBI()
        {
            if (!string.IsNullOrEmpty(UMBI.InputValue) && UMBI.InputValue.Split('/').Length == 2)
            {
                UMBI.hasCoverage = true;
                string[] limits = UMBI.InputValue.Split('/');
                limits[0] = limits[0] + "000";
                limits[1] = limits[1] + "000";
                string newInputValue = string.Format("{0}/{1}", limits[0], limits[1]);

                UMBI.InputValue = newInputValue;
                UMBI.Value1 = limits[0];
                UMBI.Value2 = limits[1];
            }
        }

        private void ProcessMed()
        {
            //TODO
        }
    }
}
