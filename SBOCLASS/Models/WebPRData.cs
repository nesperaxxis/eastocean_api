using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;

namespace SBOCLASS.Models
{
    public class WebPRHeader
    {
        public string PReqID { get; set; }
        public string PRDate { get; set; }
        public string SupplierID { get; set; }
        public double AmountTotal { get; set; }
        public string RequestBy { get; set; }
        public string Remark { get; set; }
        public virtual List<WebPRDetail> Header_Lines { get; set; }
    }
    public class WebPRDetail
    {
        public string PReqID { get; set; }
        public string ItemID { get; set; }
        public string UOMID { get; set; }
        public string GLCode { get; set; }
        public string ProjectCode { get; set; }
        public double UnitPrice { get; set; }
        public double Quantity { get; set; }
        public string AmountTotal { get; set; }
        public string Remark { get; set; }
    }
}
