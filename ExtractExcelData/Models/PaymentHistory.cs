using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractExcelData
{
    public class PaymentHistory
    {
        public string ParcelID { get; set; }
        public string LastPaymentAmount { get; set; }
        public string LastPaymentDate { get; set; }
        public string FiscalTaxYearPayments { get; set; }
        public string PriorCalendarYearPayments { get; set; }
        public string CurrentCalendarYearPayments { get; set; }
    }
}
