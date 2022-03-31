using SBOCLASS.Class;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models.EOA
{
    public class StockAdjustmentNeg
    {
        public string Ref2;
        public string PostDate;     //Date string in 'yyyyMMdd' format
        public string Remark;
        public string JournalRemark;
        public string WMSTransId;

        public List<StockAdjustmentNegDetail> Lines = new List<StockAdjustmentNegDetail>();

        public DateTime GetPostDate()
        {
            if (DateTime.TryParseExact(PostDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public bool Validate()
        {
            if (!DateTime.TryParseExact(PostDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                throw new Exception($"Invalid PostDate('{PostDate}'). Must be valid date string in the format of 'yyyyMMdd'");

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

    public class StockAdjustmentNegDetail
    {
        public int LineNo = -1;
        public string ItemCode;
        public string ItemName;
        public string UOM;
        public double Quantity = 0.0;
        public string SNBCode;
        public string Whse;
        public string BinCode;
        public string WMSTransId;
        private int? _binEntry = null;          //Marked as Null first. Will be validated later
        private string _itemManagedBy = null;   //Marked as Null first. Will be validated later
        private bool? _isItemExist = null;      //Marked as Null first. Will be validated later
        private bool? _isBinWarehouse = null;   //Marked as Null first. Will be validated later
        private string _itemInventoryUOM = null;//Marked as Null first. Will be validated later

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
            }
            else if (!String.IsNullOrWhiteSpace(BinCode))
                throw new Exception($"Line [{WMSTransId}]. Invalid Bin Code '{BinCode}'. Warehouse '{Whse}' not managed by bin.");

            if (!IsItemExists(company)) throw new Exception($"Line [{WMSTransId}]. ItemCode '{ItemCode}' does not exists.");
            _itemManagedBy = IsItemManagedBySnB(company);
            if (!String.IsNullOrWhiteSpace(_itemManagedBy))
            {
                //Serial and batch must exists for Issue type of documents
                if (String.IsNullOrWhiteSpace(SNBCode)) throw new Exception($"Line [{WMSTransId}]. Item is managed by {(_itemManagedBy == "S" ? "Serial" : "Batch")}. SnBCode must be provided.");
                if (!IsValidSnb(company)) throw new Exception($"Line [{WMSTransId}]. Item Serial/Batch '{SNBCode}' does not exist or not enough stock.");
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


    }
}
