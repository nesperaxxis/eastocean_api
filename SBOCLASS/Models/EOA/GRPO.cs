using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SBOCLASS.Class;

namespace SBOCLASS.Models.EOA
{
    public class GRPO
    {
        public string CardCode="";
        public string PostDate="";     //Date string in 'yyyyMMdd' format
        public string Remark= "";
        //public int SlpCode = 0; // change to accept the SlpName
        public string SlpCode = "";
        public string WMSTransId = "";
        private string _sapCardCode = null;     //Marked as Null first. Will be validated later.

        public List<GRPODetail> Lines = new List<GRPODetail>();

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

            if (String.IsNullOrWhiteSpace(CardCode))
                throw new Exception($"CardCode is missing");

            var validate = Lines.Select(x => x.Validate()).ToList();

            return true;
        }

        public string GetSAPCardCode(SAPbobsCOM.Company company)
        {
            if (_sapCardCode != null) return _sapCardCode;

            if (!SBOSupport.IsCardCodeExists(company, CardCode, out string result))
                throw new Exception(result);

            _sapCardCode = result;
            return _sapCardCode;
        }

        public string GetSAPSlpCode(SAPbobsCOM.Company company)
        {
            if (_sapCardCode != null) return _sapCardCode;

            if (!SBOSupport.IsSlpCodeExists(company, SlpCode, out string result))
                throw new Exception(result);

            _sapCardCode = result;
            return _sapCardCode;
        }

        public bool ValidateLine(SAPbobsCOM.Company company)
        {
            var validate = Lines.Select(x => x.Validate(company)).ToList();
            return true;

        }
    }

    public class GRPODetail
    {
        public string BaseType = "";
        public int BaseKey = 0;
        public int BaseLine = -1;
        public string ItemCode = "";
        public string ItemName = "";
        public string UOM = "";
        public double Quantity = 0.0;
        public string SNBCode = "";
        public string Whse = "";
        public string BinCode = "";
        public string WMSTransId = "";
        private int? _binEntry = null;          //Marked as Null first. Will be validated later
        private string _itemManagedBy = null;   //Marked as Null first. Will be validated later
        private bool? _isItemExist = null;      //Marked as Null first. Will be validated later
        private bool? _isBinWarehouse = null;   //Marked as Null first. Will be validated later
        private bool? _useBaseUnit = null;      //Merked as null first. Will be validated later
        private string _itemInventoryUOM = null;//Marked as Null first. Will be validated later
        private string _itemPurchaseUOM = null; //Marked as Null first. Will be validated later
        private double _numInBuy = 0.0;
        private double _qtyUOMConversion = 0.0;

        public bool Validate()
        {
            if (String.IsNullOrWhiteSpace(WMSTransId)) throw new Exception($"Each line must have a unique WMSTransId.");
            if (String.IsNullOrWhiteSpace(BaseType)) throw new Exception($"Line {WMSTransId}. Must have BaseType.");
            if (BaseType != "22" && !String.IsNullOrWhiteSpace(BaseType)) throw new Exception($"Line {WMSTransId}. Invalid BaseType. Valid values are: 22 - Purchase Order");
            if (BaseKey == 0) throw new Exception($"Line {WMSTransId}. BaseKey must be provided.");
            if (BaseLine <= -1) throw new Exception($"Line {WMSTransId}. BaseLine must be provided.");
            if (String.IsNullOrWhiteSpace(ItemCode)) throw new Exception($"Line {WMSTransId}. ItemCode is missing");
            if (String.IsNullOrWhiteSpace(Whse)) throw new Exception($"Line {WMSTransId}. Whse code is missing");
            if (Quantity <= 0.0) throw new Exception($"Line {WMSTransId}. Quantity must be greater than 0.0");
            UOM = UOM ?? "";
            return true;
        }

        public bool Validate(SAPbobsCOM.Company company)
        {
            //If base type is provided, base entry and base line must exists in SAP
            string result = "";
            if (!String.IsNullOrWhiteSpace(BaseType))
            {
                if (!IsBaseEntryExists(company, out result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
                if (!IsBaseLineExists(company, out result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
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
            if (!String.IsNullOrWhiteSpace(_itemManagedBy))
            {
                if (String.IsNullOrWhiteSpace(SNBCode)) throw new Exception($"Line [{WMSTransId}]. Item is managed by {(_itemManagedBy == "S" ? "Serial" : "Batch")}. SnBCode must be provided.");
                //Serial and batch must exists for documents of issuing only.
                //if (!IsValidSnb(company)) throw new Exception($"Line [{WMSTransId}]. Item Serial/Batch '{SNBCode}' does not exist or not enough stock.");
            }

            GetItemNumInBuy(company);

            if (!IsInventoryUOM(company, out _) && !IsPurchaseUOM(company, out _)) throw new Exception($"Line [{WMSTransId}]. UOM must be either Inventory UOM '{_itemInventoryUOM}' or Purchase UOM '{_itemPurchaseUOM}'.");

            return true;

        }

        private bool IsBaseEntryExists(SAPbobsCOM.Company company, out string result)
        {
            bool exists = true;
            result = "";
            string baseTable = String.IsNullOrWhiteSpace(BaseType) ? "" : SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)GetSAPBaseType());
            if (baseTable != "")
            {
                string sql = $"SELECT ISNULL(MAX(CASE WHEN CANCELED <> 'N' THEN 'C' ELSE \"InvntSttus\" END),'')  FROM O{baseTable} WHERE \"DocEntry\" = {BaseKey} ";
                string docStatus = SBOSupport.GETSINGLEVALUE(sql, company);
                if (String.IsNullOrWhiteSpace(docStatus))
                {
                    exists = false;
                    result = $"BaseKey {BaseKey}  does not exists. [O{baseTable}].";
                }
                else if (docStatus == "C")
                {
                    exists = false;
                    result = $"Base document status is 'Closed'/'Canceled'. [O{baseTable}]{BaseKey}";
                }
            }

            return exists;
        }

        private bool IsBaseLineExists(SAPbobsCOM.Company company, out string result)
        {
            bool exists = IsBaseEntryExists(company, out result);
            if (!exists) return false;

            string baseTable = String.IsNullOrWhiteSpace(BaseType) ? "" : SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)GetSAPBaseType());
            if (baseTable != "")
            {
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
            }

            return exists;
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

        public bool IsPurchaseUOM(SAPbobsCOM.Company company, out string result)
        {
            result = "";
            if (_itemPurchaseUOM == null)
                _itemPurchaseUOM = SBOSupport.GetItemPurchaseUOM(company, ItemCode);

            if (_itemPurchaseUOM.ToUpper().Trim() != UOM.ToUpper().Trim() )
            {
                result = $"UOM must be Item Purchase UOM ({_itemPurchaseUOM})";
                return false;
            }

            return  true;

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

        public int GetSAPBaseType()
        {
            if (BaseType == "22") return 22;
            if (BaseType == "") return -1;
            throw new Exception("Invalid base type. Valid Values are 22 - Purchase Order.");
        }

        public double GetItemNumInBuy(SAPbobsCOM.Company company)
        {
            if (_numInBuy != 0.0) return _numInBuy;

            if (GetSAPBaseType() != -1)
            {
                double numInBuy =SBOSupport.GetItemNumPerMsr(company, GetSAPBaseType(), BaseKey, BaseLine, out string baseDocUom, out bool useBaseUnit);
                if (baseDocUom.ToUpper() == UOM)    //If the UOM is equal to Base document UOM, Get them numInBuy from Document Line numpermsr
                {
                    _numInBuy = numInBuy;
                    _useBaseUnit = useBaseUnit;
                }
                else if (IsPurchaseUOM(company, out _)) //If the UOM is not from Base document UOM, and is ItemMasterData Purchase UOM, Get them numInBuy from Item Master Data NumInBUy
                {
                    _numInBuy = SBOSupport.GetItemMasterDataNumInBuy(company, ItemCode);
                    _useBaseUnit = false;
                }
                else if (IsInventoryUOM(company, out _)) //If the UOM is not from Base document UOM, and is ItemMasterData Inventory UOM, the numinbuy = 1
                {
                    _numInBuy = 1;
                    _useBaseUnit = true;
                }
                else
                    throw new Exception("Invalid UOM. UOM must be either the purchase UOM or base inventory UOM");
            }
            else
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
