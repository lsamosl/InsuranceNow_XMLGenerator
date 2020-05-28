using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class General
    {
        public string PolicyNumber { get; set; }
        public string MedCover { get; set; }
        public string BiCover { get; set; }
        public string Umbi { get; set; }
        public string UmpdCdw { get; set; }
        public string LimMex { get; set; }
        public string RoadAssis { get; set; }
        public string PdCover { get; set; }
        public GeneralCoverages Coverages { get; set; }
    }
}
