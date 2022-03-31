using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;


namespace SBOCLASS.Models
{
    public class MAGENTO_ARINVOICE
    {
        public string OrderID { get; set; }
        public string OrderStatus { get; set; }
        public double CartBefPriceDisc { get; set; }
        public double CXACartDisc { get; set; }
        public double GrandTotal_Inc { get; set; }
        public string EmployeeName { get; set; }
        public string EmployeeId { get; set; }
        public string EmailAddress { get; set; }
        public string FullName { get; set; }
        public DateTime PurchaseDateSGT { get; set; }
        public DateTime CompletionDateSGT { get; set; }
        public string Currency { get; set; }
        public string ModeOfPayment { get; set; }
        public double AdminFeeExGST { get; set; }
        public double AdminFeeGST { get; set; }
        public string StripeID { get; set; }
        public string StripeStatus { get; set; }
        public virtual List<MAGENTO_VOUCHER_SHIPMENT> Voucher_SHIPMENT { get; set; }
    }
    public class MAGENTO_VOUCHER_SHIPMENT
    {
        public string OrderID { get; set; }
        public string Type{ get; set; }
        public string Status { get; set; }
        public string SubOrderId { get; set; }
        public string VoucherID { get; set; }
        public string UniqueID { get; set; }
        public double Quantity { get; set; }
        public string VoucherNumber { get; set; }
        public string ShipmentID { get; set; }
        public string ProviderEntityName { get; set; }
        public string ProviderEntityId { get; set; }
        public string ProductId { get; set; }
        public string ProductName { get; set; }
        public DateTime PurchaseDateSGT { get; set; }
        public DateTime CompletionDateSGT { get; set; }
        public DateTime ExpirationDateSGT { get; set; }
        public string Currency { get; set; }
        public double ProviderListedPriceGST { get; set; }
        public double ProviderListedPriceExGST { get; set; }
        public double DisplayPriceGST { get; set; }
        public double DisplayPriceExGST { get; set; }
        public double CommissionRate { get; set; }
        public double CommissionAmountGST { get; set; }
        public double CommissionAmountexGST { get; set; }
        public double CXAProductDiscount { get; set; }
        public string GstType { get; set; }
        public string MRV_NMRV { get; set; }
        public string PAYMENTTYPE { get; set; }
      

    }

}
