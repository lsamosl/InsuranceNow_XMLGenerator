using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services.Models
{
    public class ResponseDetokenize
    {
        public bool Result { get; set; }
        public string DetokenizedString { get; set; }
        public string Message { get; set; }
    }
}
