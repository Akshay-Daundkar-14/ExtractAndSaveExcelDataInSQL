using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractExcelData.Models
{
    public class ParcelRecord
    {
        public string ParcelNumber { get; set; }
        public string ParNum { get; set; }
        public string DocName { get; set; }
        public string DocName1 { get; set; }
        public string Status { get; set; }
        public string Remark { get; set; }
        public string StartTime { get; set;}
        public string EndTime { get; set; }

        public string TotalTime { get; set; }

        
    }

}
