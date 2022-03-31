using SBOCLASS.Class;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models.EOA
{
    public class PickList
    {
        public string PickDate;     //yyyyMMdd
        public int User = 0;
        public string Remark;
        public string WMSTransId;
        private string _baseCardCode;       //CardCodes of the base lines.
        private string _sapCardCode;        //CardCode as in SAP after validation of _baseCardCode.

        public List<PickListDetail> Lines = new List<PickListDetail>();

        public DateTime GetPickDate()
        {
            if (DateTime.TryParseExact(PickDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public bool Validate()
        {
            if (!DateTime.TryParseExact(PickDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                throw new Exception($"Invalid PickDate('{PickDate}'). Must be valid date string in the format of 'yyyyMMdd'");

            if (String.IsNullOrWhiteSpace(WMSTransId))
                throw new Exception($"WMSTransId is missing");

            var validate = Lines.Select(x => x.Validate()).ToList();

            return true;
        }

        public bool ValidateLine(SAPbobsCOM.Company company)
        {
            if (Lines.Count == 0) throw new Exception("Pick list must have at least 1 line.");
            var validate = Lines.Select(x => x.Validate(company)).ToList();

            var baseType = Lines.Select(x => x.BaseType).Distinct().ToList();
            if (baseType.Count > 1) throw new Exception("PickList cannot mix line base type.");
            if(baseType.Count == 1 && baseType[0] == "13R")
            {
                //based on AR Reserve invoice - must only have 1 cardCode.
                var cardCodes = Lines.Select(x => x.CardCode).Distinct().ToList();
                if (cardCodes.Count > 1) throw new Exception("PickList cannot mix line card code");
                _baseCardCode = cardCodes[0];
            }
            return true;

        }

        public string GetSAPCardCode(SAPbobsCOM.Company company)
        {
            if (_sapCardCode != null) return _sapCardCode;
            if (_baseCardCode == null) ValidateLine(company);

            if (_baseCardCode == "")
            {
                _sapCardCode = "";
                return _sapCardCode;
            }

            SBOSupport.IsCardCodeExists(company, _baseCardCode, out _sapCardCode);

            return _sapCardCode;

        }

    }

    public class PickListDetail
    {
        public int BaseKey = 0;
        public int BaseNum = 0;
        public string BaseType="";
        public string PostDate="";
        public string CardCode = "";
        public string NumAtCard = "";
        public string DeliveryDate = "";     //Date string in 'yyyyMMdd'
        public int BaseLine = -1;
        public string ItemCode = "";
        public string ItemName = "";
        public string Whse = "";
        public string BinCode = "";
        public string SNBCode = "";
        public string UOM = "";
        public Double PickedQty = 0.0;
        public string WMSTransId = "";
        private int? _binEntry = null;          //Marked as Null first.
        private string _itemManagedBy = null;   //Marked as Null first. 
        private bool? _isItemExist = null;      //Marked as Null first.
        private bool? _isBinWarehouse = null;   //Marked as Null first.
        private string _itemInventoryUOM = null;//Marked as Null first. Will be validated later
        private string _itemSaleUOM = null; //Marked as Null first. Will be validated later
        private double _numInSale = 0.0;


        public DateTime GetDeliveryDate()
        {
            if (DateTime.TryParseExact(DeliveryDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public bool Validate()
        {
            if (String.IsNullOrWhiteSpace(WMSTransId)) throw new Exception($"Each line must have a unique WMSTransId.");
            if (String.IsNullOrWhiteSpace(BaseType)) throw new Exception($"Line [{WMSTransId}]. Must have BaseType.");
            if (BaseType != "17" && BaseType != "13R") throw new Exception($"Line [{WMSTransId}]. Invalid BaseType. Valid values are: 17 - Sales Order, 13R - Reserve Invoice.");
            if (BaseKey == 0) throw new Exception($"Line [{WMSTransId}]. BaseKey must be provided.");
            if (BaseLine <= -1) throw new Exception($"Line [{WMSTransId}]. BaseLine must be provided.");
            if (String.IsNullOrWhiteSpace(CardCode)) throw new Exception($"Line [{WMSTransId}]. CardCode is missing");
            if (String.IsNullOrWhiteSpace(DeliveryDate)) throw new Exception($"Line [{WMSTransId}]. DeliveryDate is missing");
            if (GetDeliveryDate() == DateTime.FromOADate(0)) throw new Exception($"Line [{WMSTransId}]. Invalid DeliveryDate. Must be in the format of 'yyyyMMdd'");
            if (String.IsNullOrWhiteSpace(ItemCode)) throw new Exception($"Line [{WMSTransId}]. ItemCode is missing");
            if (String.IsNullOrWhiteSpace(Whse)) throw new Exception($"Line [{WMSTransId}]. Whse code is missing");
            UOM = UOM ?? "";

            return true;
        }

        public bool Validate(SAPbobsCOM.Company company)
        {
            //If base type is provided, base entry and base line must exists in SAP
            string result = "";
            if (!String.IsNullOrWhiteSpace(BaseType))
            {
                if (!IsBaseEntryExists(company, out  result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
                if (!IsBaseLineExists(company, out  result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
            }

            if (!IsWarehouseExists(company, out result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
            if (IsBinWarehouse(company))
            {
                //Bin warehouse, line must have bin code;
                if (String.IsNullOrWhiteSpace(BinCode)) throw new Exception($"Line [{WMSTransId}]. Bin managed warehouse must have bin code.");
                if (GetBinEntry(company) == 0) throw new Exception($"Line [{WMSTransId}]. Invalid Bin Code '{BinCode}'. Bin does not exists in Warehouse '{Whse}'.");
            }

            if (!IsItemExists(company)) throw new Exception($"Line [{WMSTransId}]. ItemCode '{ItemCode}' does not exists.");
            _itemManagedBy = IsItemManagedBySnB(company);
            if(!String.IsNullOrWhiteSpace(_itemManagedBy))
            {
                //Serial and batch must exists for delivery
                if (String.IsNullOrWhiteSpace(SNBCode)) throw new Exception($"Line [{WMSTransId}]. Item is managed by {(_itemManagedBy == "S" ? "Serial" : "Batch")}. SnBCode must be provided.");
                if (!IsValidSnb(company)) throw new Exception($"Line [{WMSTransId}]. Item Serial/Batch '{SNBCode}' does not exist or not enough stock.");
            }

            if (!IsInventoryUOM(company, out result)) throw new Exception($"Line [{WMSTransId}]. {result}");

            return true;

        }

        private bool IsBaseEntryExists(SAPbobsCOM.Company company, out string result)
        {
            bool exists = true;
            result = "";
            string baseTable = BaseType == "17" ? "RDR" : "INV";
            string sql = $"SELECT ISNULL(MAX(CASE WHEN CANCELED <> 'N' THEN 'C' ELSE \"InvntSttus\" END),'')  FROM O{baseTable} WHERE \"DocEntry\" = {BaseKey} ";
            string docStatus = SBOSupport.GETSINGLEVALUE(sql, company);
            if(String.IsNullOrWhiteSpace(docStatus))
            {
                exists = false;
                result = $"BaseKey {BaseKey}  does not exists. [O{baseTable}].";
            } else if(docStatus == "C")
            {
                exists = false;
                result = $"Base document status is 'Closed'/'Canceled'. [O{baseTable}]{BaseKey}";
            }

            return exists;
        }

        private bool IsBaseLineExists(SAPbobsCOM.Company company, out string result)
        {
            bool exists = IsBaseEntryExists(company, out result);
            if (!exists) return false;

            string baseTable = BaseType == "17" ? "RDR" : "INV";
            string sql = $"SELECT ISNULL(MAX(\"LineStatus\"),'') FROM  {baseTable}1 WHERE \"DocEntry\" = {BaseKey} AND \"LineNum\" = {BaseLine} ";
            string docStatus = SBOSupport.GETSINGLEVALUE(sql, company);
            if (String.IsNullOrWhiteSpace(docStatus))
            {
                exists = false;
                result = $"BaseKey\\BaseLine {BaseKey}\\{BaseLine}  does not exists. [{baseTable}1].";
            }
            else if (docStatus == "C")
            {
                exists = false;
                result = $"Base Line status is 'Closed'. [{baseTable}1]{BaseKey}\\{BaseLine}";
            }

            return exists;
        }

        private bool IsWarehouseExists(SAPbobsCOM.Company company, out string result)
        {
            return SBOSupport.IsWarehouseExists(company, Whse, out result);
        }

        public bool IsInventoryUOM(SAPbobsCOM.Company company, out string   result)
        {
            result = "";
            if (_itemInventoryUOM == null)
                _itemInventoryUOM = SBOSupport.GetItemInventoryUOM(company, ItemCode);

            if( _itemInventoryUOM.ToUpper().Trim() != UOM.ToUpper().Trim())
            {
                result = $"UOM must be Item Inventory UOM ({_itemInventoryUOM})";
                return false;
            }

            return true;
            
        }

        public int GetBinEntry(SAPbobsCOM.Company company)
        {
            if (_binEntry !=null)
                return _binEntry??0;

            if (String.IsNullOrWhiteSpace(BinCode))
            {
                _binEntry = 0;
                return _binEntry??0;
            }

            _binEntry = SBOSupport.GetBinEntry(company, BinCode, out string result);
            if(_binEntry == 0) throw new Exception($"Line {WMSTransId}. {result}");
            return _binEntry??0;
        }

        public bool IsItemExists(SAPbobsCOM.Company company)
        {
            if (_isItemExist != null) return _isItemExist??true;
            if (String.IsNullOrWhiteSpace(ItemCode))
            {
                _isItemExist = false;
                return _isItemExist??false;
            }

            _isItemExist = SBOSupport.IsItemExists(company, ItemCode);
            return _isItemExist??true;
        }

        public string IsItemManagedBySnB(SAPbobsCOM.Company company)
        {
            if (_itemManagedBy != null)
                return _itemManagedBy;
            if (String.IsNullOrWhiteSpace(ItemCode))
            {
                _itemManagedBy = "";
                return _itemManagedBy??"";
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

        public int GetSAPBaseType()
        {
            if (BaseType == "17") return 17;
            if (BaseType == "13R") return 13;
            throw new Exception("Invalid base type. Valid Values are 17 - Sales Order, 13R - Reserve Invoice.");
        }

        public bool IsSalesUOM(SAPbobsCOM.Company company, out string result)
        {
            result = "";
            if (_itemSaleUOM == null)
                _itemSaleUOM = SBOSupport.GetItemSalesUOM(company, ItemCode);

            if (_itemSaleUOM.ToUpper().Trim() != UOM.ToUpper().Trim())
            {
                result = $"UOM must be Item Sales UOM ({_itemSaleUOM})";
                return false;
            }

            return true;

        }

        public double GetItemNumInSale(SAPbobsCOM.Company company)
        {
            if (_numInSale != 0.0) return _numInSale;

            _numInSale = SBOSupport.GetItemNumPerMsr(company, GetSAPBaseType(), BaseKey, BaseLine, out string baseDocUom, out bool useBaseUnit  );
            if (_numInSale == 0) _numInSale = 1.0;

            return _numInSale;
        }
    }
}
