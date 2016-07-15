using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddressParser
{
    public class StreetAddress
    {
        public string HouseNumber { get; set; }
        public string StreetPrefix { get; set; }
        public string StreetName { get; set; }
        public string StreetType { get; set; }
        public string StreetSuffix { get; set; }
        public string Apt { get; set; }
    }
}
