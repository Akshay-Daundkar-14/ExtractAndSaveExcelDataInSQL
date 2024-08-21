using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractExcelData
{
   public class PaymentData
    {
        public string ParcelNumber { get; set; }
        public string BatchNumber { get; set; }
        public string PaymentDate { get; set; }
        public string InterestDate { get; set; }
        public string Payee { get; set; }
        public string BatchAmount { get; set; }
    }
}
