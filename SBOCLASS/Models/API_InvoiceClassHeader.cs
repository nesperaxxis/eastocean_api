using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models
{
        public class API_InvoiceClassHeader
    {
        public string U_TransId { get; set; }
        public string CardCode { get; set; }
        public DateTime DocDate { get; set; }
        public string NumAtCard { get; set; }
        public virtual List<API_InvoiceClassDetails> Details { get; set; }
    }
    public class API_InvoiceClassDetails
    {
        public string ItemCode { get; set; }
        public double Quantity { get; set; }
        public double Price { get; set; }
    }

    public class GetPaymentStatus
    {
        public string U_TransId { get; set; }
        public string Status { get; set; }
        public double DocTotal { get; set; }
        public double AppliedAmount { get; set; }
        public double BalanceDue { get; set; }
        public string PaidDate { get; set; }
        public string Currency { get; set; }
    }

   
}
