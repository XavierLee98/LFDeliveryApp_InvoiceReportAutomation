using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceReportAutomation.Model
{
    public class RequestModel
    {
        public int Id { get; set; }

        public string CompanyId { get; set; }

        public int InvoiceDocEntry { get; set; }

        public int IsTry { get; set; }

        public bool? IsCreated { get; set; }

        public string FilePath { get; set; }

        public DateTime? CreatedDate { get; set; }
    }
}
