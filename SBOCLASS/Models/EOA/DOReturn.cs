using SBOCLASS.Class;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models.EOA
{
    public class DOReturn
    {
        public string CardCode="";
        public string PostDate="";         //Date string in 'yyyyMMdd' format
        public string DeliveryDate="";     //Date string in 'yyyyMMdd' format
        public string Remark="";
        public string DocType = "RETURNS";
        public string WMSTransId="";


        public List<DOReturnDetail> Lines = new List<DOReturnDetail>();

        public DateTime GetPostDate()
        {
            if (DateTime.TryParseExact(PostDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public DateTime GetDeliveryDate()
        {
            if (DateTime.TryParseExact(DeliveryDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public bool Validate()
        {
            if (String.IsNullOrWhiteSpace(CardCode))
                throw new Exception($"Customer Code is missing");

            if (!DateTime.TryParseExact(PostDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out _))
                throw new Exception($"Invalid PostDate('{PostDate}'). Must be valid date string in the format of 'yyyyMMdd'");

            if (!DateTime.TryParseExact(DeliveryDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out _))
                throw new Exception($"Invalid DeliveryDate('{PostDate}'). Must be valid date string in the format of 'yyyyMMdd'");

            if (String.IsNullOrWhiteSpace(WMSTransId))
                throw new Exception($"WMSTransId is missing");

            var validate = Lines.Select(x => x.Validate()).ToList();

            return true;
        }

        public bool ValidateLine(SAPbobsCOM.Company company)
        {
            var validate = Lines.Select(x => x.Validate(company)).ToList();
            return true;

        }
    }

    public class DOReturnDetail
    {
        public int LineNo = -1;
        public string ItemCode;
        public string ItemName;
        public string UOM;
        public double Quantity = 0.0;
        public string SNBCode;
        public string Whse;
        public string BinCode;
        public string ReturnReason;
        public string WMSTransId;
        private int? _binEntry = null;          //Marked as Null first. Will be validated later
        private string _itemManagedBy = null;   //Marked as Null first. Will be validated later
        private bool? _isItemExist = null;      //Marked as Null first. Will be validated later
        private bool? _isBinWarehouse = null;   //Marked as Null first. Will be validated later
        private string _itemInventoryUOM = null;//Marked as Null first. Will be validated later
        private string _itemPurchaseUOM = null; //Marked as Null first. Will be validated later
        private double _numInBuy = 0.0;
        private double _qtyUOMConversion = 0.0;

        public bool Validate()
        {
            if (String.IsNullOrWhiteSpace(WMSTransId)) throw new Exception($"Each line must have a unique WMSTransId.");
            if (String.IsNullOrWhiteSpace(ItemCode)) throw new Exception($"Line {WMSTransId}. ItemCode is missing");
            if (String.IsNullOrWhiteSpace(Whse)) throw new Exception($"Line {WMSTransId}. Whse code is missing");
            if (Quantity<=0.0) throw new Exception($"Line {WMSTransId}. Quantity must be greater than 0.0");
            UOM = UOM ?? "";

            return true;
        }

        public bool Validate(SAPbobsCOM.Company company)
        {
            string result = "";

            if (!IsWarehouseExists(company, out result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
            if (IsBinWarehouse(company))
            {
                //Bin warehouse, line must have bin code;
                if (String.IsNullOrWhiteSpace(BinCode)) throw new Exception($"Line [{WMSTransId}]. Bin managed warehouse must have bin code.");
                if (GetBinEntry(company) == 0) throw new Exception($"Line [{WMSTransId}]. Invalid Bin Code '{BinCode}'. Bin does not exists in Warehouse '{Whse}'.");
            } else if (!String.IsNullOrWhiteSpace(BinCode))
                throw new Exception($"Line [{WMSTransId}]. Invalid Bin Code '{BinCode}'. Warehouse '{Whse}' not managed by bin.");

            if (!IsItemExists(company)) throw new Exception($"Line [{WMSTransId}]. ItemCode '{ItemCode}' does not exists.");
            _itemManagedBy = IsItemManagedBySnB(company);
            if (!String.IsNullOrWhiteSpace(_itemManagedBy))
            {
                if (String.IsNullOrWhiteSpace(SNBCode)) throw new Exception($"Line [{WMSTransId}]. Item is managed by {(_itemManagedBy == "S" ? "Serial" : "Batch")}. SnBCode must be provided.");
            }

            if (!IsInventoryUOM(company, out result)) throw new Exception($"Line [{WMSTransId}]. {result}");

            return true;

        }

        private bool IsWarehouseExists(SAPbobsCOM.Company company, out string result)
        {
            return SBOSupport.IsWarehouseExists(company, Whse, out result);
        }

        public bool IsInventoryUOM(SAPbobsCOM.Company company, out string result)
        {
            result = "";
            if (_itemInventoryUOM == null)
                _itemInventoryUOM = SBOSupport.GetItemInventoryUOM(company, ItemCode);

            if (_itemInventoryUOM.ToUpper().Trim() != UOM.ToUpper().Trim())
            {
                result = $"UOM must be Item Inventory UOM ({_itemInventoryUOM})";
                return false;
            }

            return true;
        }

        public int GetBinEntry(SAPbobsCOM.Company company)
        {
            if (_binEntry != null)
                return _binEntry ?? 0;

            if (String.IsNullOrWhiteSpace(BinCode))
            {
                _binEntry = 0;
                return _binEntry ?? 0;
            }

            _binEntry = SBOSupport.GetBinEntry(company, BinCode, out string result);
            if (_binEntry == 0) throw new Exception($"Line {WMSTransId}. {result}");
            return _binEntry ?? 0;
        }

        public bool IsItemExists(SAPbobsCOM.Company company)
        {
            if (_isItemExist != null) return _isItemExist ?? true;
            if (String.IsNullOrWhiteSpace(ItemCode))
            {
                _isItemExist = false;
                return _isItemExist ?? false;
            }

            _isItemExist = SBOSupport.IsItemExists(company, ItemCode);
            return _isItemExist ?? true;
        }

        public string IsItemManagedBySnB(SAPbobsCOM.Company company)
        {
            if (_itemManagedBy != null)
                return _itemManagedBy;
            if (String.IsNullOrWhiteSpace(ItemCode))
            {
                _itemManagedBy = "";
                return _itemManagedBy ?? "";
            }

            _itemManagedBy = SBOSupport.IsItemManagedBySnB(company, ItemCode);

            return _itemManagedBy ?? "";
        }

        public bool IsBinWarehouse(SAPbobsCOM.Company company)
        {
            if (_isBinWarehouse != null) return _isBinWarehouse ?? true;
            if (String.IsNullOrWhiteSpace(Whse))
            {
                _isBinWarehouse = false;
                return _isBinWarehouse ?? false;
            }

            _isBinWarehouse = SBOSupport.IsBinWarehouse(company, Whse);
            return _isBinWarehouse ?? true;
        }

        public bool IsValidSnb(SAPbobsCOM.Company company)
        {
            return SBOSupport.IsSnBExists(company, ItemCode, SNBCode, out string result);
        }

        public double GetItemNumInBuy(SAPbobsCOM.Company company)
        {
            if (_numInBuy != 0.0) return _numInBuy;

            _numInBuy = SBOSupport.GetItemMasterDataNumInBuy(company, ItemCode);
            if (_numInBuy == 0) _numInBuy = 1.0;

            return _numInBuy;
        }

        public double GetUomQtyConversion(SAPbobsCOM.Company company)
        {
            if (_qtyUOMConversion != 0.0) return _qtyUOMConversion;

            if (IsInventoryUOM(company, out _))
            {
                _qtyUOMConversion = 1.0;
                return _qtyUOMConversion;
            }
            else
            {
                _qtyUOMConversion = GetItemNumInBuy(company);
                return _qtyUOMConversion;
            }
        }
    }
}
