using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

namespace SBOCLASS.Models
{
    public class HEADER
    {
        public string InvNo { get; set; }
        public string CompanyCode { get; set; }
        public string InvType { get; set; }
        public string PONO { get; set; }
        public string PODescription { get; set; }
        public string VendorCode { get; set; }
        public string ProjectCode { get; set; }
        public string CostCenter { get; set; }
        public string Currency { get; set; }
        public DateTime PostingDate { get; set; }
        public string FinanceAccount { get; set; }
        public virtual List<DETAILS> Header_Lines { get; set; }

    }
    public class DETAILS
    {
        public string InvNo { get; set; }
        public int LineNo { get; set; }
        public string FinAcc { get; set; }
        public decimal Amount { get; set; }
        public string Comments { get; set; }
        public string VATName { get; set; }
        public decimal VATAmount { get; set; }
        
    }

    public class INVOICEHEADER
    {
        public string InvNo { get; set; }
        public string CompanyCode { get; set; }
        public string InvType { get; set; }
        public string PONO { get; set; }
        public string InvStatus { get; set; }
    }
    public class ResponseResult
    {
        public string RecordStatus { get; set; }
        public string ErrorDescription { get; set; }
    }

}