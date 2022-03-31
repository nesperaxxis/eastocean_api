using SBOCLASS.Class;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models.EOA
{
    public class InventoryTransfer 
    {
        public string InTransitCode;
        public string ReceiptDate;  //Date string in 'yyyyMMdd' format
        public string ShipmentDate; //Date string in 'yyyyMMdd' format
        public string PostDate;     //Date string in 'yyyyMMdd' format
        public string Remark;
        public string JournalRemark;
        public string WMSTransId;
        public string RqWMSTransId;     //To Identify the base Request Document, if copied from.

        public List<InventoryTransferDetail> Lines = new List<InventoryTransferDetail>();

        public DateTime GetPostDate()
        {
            if (DateTime.TryParseExact(PostDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public DateTime GetReceiptDate()
        {
            if (DateTime.TryParseExact(ReceiptDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public DateTime GetShipmentDate()
        {
            if (DateTime.TryParseExact(ShipmentDate, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;
            else
                return DateTime.FromOADate(0);
        }

        public bool Validate()
        {
            if (GetPostDate()==DateTime.FromOADate(0))
                throw new Exception($"Invalid PostDate('{PostDate}'). Must be valid date string in the format of 'yyyyMMdd'");

            if (String.IsNullOrWhiteSpace(WMSTransId))
                throw new Exception($"WMSTransId is missing");

            var validate = Lines.Select(x => x.Validate()).ToList();
           
            return true;
        }

        public bool ValidateLine(SAPbobsCOM.Company company)
        {
            var validate = Lines.Select(x => x.Validate(company, RqWMSTransId)).ToList();
            return true;

        }
    }

    public class InventoryTransferDetail
    {
        public string ItemCode;
        public string ItemName;
        public int LineNo = -1;
        public string UOM;
        public double Quantity = 0.0;
        public string SNBCode;
        public string FromWhse;
        public string FromBin;
        public string ToWhse;
        public string ToBin;
        public string WMSTransId;
        public string RqWMSTransId;
        private string _itemManagedBy = null;   //Marked as Null first. Will be validated later
        private bool? _isItemExist = null;      //Marked as Null first. Will be validated later
        private int? _fromBinEntry = null;          //Marked as Null first. Will be validated later
        private bool? _isFromBinWarehouse = null;   //Marked as Null first. Will be validated later
        private int? _toBinEntry = null;          //Marked as Null first. Will be validated later
        private bool? _isToBinWarehouse = null;   //Marked as Null first. Will be validated later
        private string _itemInventoryUOM = null;//Marked as Null first. Will be validated later
        private int? _baseWtqEntry = null;      //Marked as Null first. Will be validated later
        private int? _baseWtqLine = null;       //Marked as Null first. Will be validated later

        public bool Validate()
        {
            if (String.IsNullOrWhiteSpace(WMSTransId)) throw new Exception($"Each line must have a unique WMSTransId.");
            if (String.IsNullOrWhiteSpace(ItemCode)) throw new Exception($"Line {WMSTransId}. ItemCode is missing");
            if (String.IsNullOrWhiteSpace(FromWhse)) throw new Exception($"Line {WMSTransId}. FromWhse code is missing");
            if (String.IsNullOrWhiteSpace(ToWhse)) throw new Exception($"Line {WMSTransId}. ToWhse code is missing");
            if (Quantity<=0.0) throw new Exception($"Line {WMSTransId}. Quantity must be greater than 0.0");
            UOM = UOM ?? "";
            if (String.IsNullOrWhiteSpace(RqWMSTransId))
            {
                _baseWtqEntry = 0;
                _baseWtqLine = 0;
            }

            return true;
        }

        public bool Validate(SAPbobsCOM.Company company, string headerWMSTransId)
        {
            //If base type is provided, base entry and base line must exists in SAP
            string result = "";
            if (!String.IsNullOrWhiteSpace(RqWMSTransId))
            {
                if (!IsBaseEntryExists(company, headerWMSTransId, out _, out _, out result)) throw new Exception(result);
            }

            //From Whs and Bin
            if (!IsWarehouseExists(company, FromWhse, out result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
            if (IsFromBinWarehouse(company))
            {
                //Bin warehouse, line must have bin code;
                if (String.IsNullOrWhiteSpace(FromBin)) throw new Exception($"Line [{WMSTransId}]. Bin managed warehouse must have bin code. [From Bin]");
                if (GetFromBinEntry(company) == 0) throw new Exception($"Line [{WMSTransId}]. Invalid 'From' Bin Code '{FromBin}'. Bin does not exists in Warehouse '{FromWhse}'.");
            }
            else if (!String.IsNullOrWhiteSpace(FromBin))
                throw new Exception($"Line [{WMSTransId}]. Invalid 'From' Bin Code '{FromBin}'. Warehouse '{FromWhse}' not managed by bin.");

            //To Whs and Bin
            if (!IsWarehouseExists(company, ToWhse, out result)) throw new Exception($"Line [{ WMSTransId }]. {result}");
            if (IsToBinWarehouse(company))
            {
                //Bin warehouse, line must have bin code;
                if (String.IsNullOrWhiteSpace(ToBin)) throw new Exception($"Line [{WMSTransId}]. Bin managed warehouse must have bin code. [To Bin]");
                if (GetToBinEntry(company) == 0) throw new Exception($"Line [{WMSTransId}]. Invalid 'To' Bin Code '{ToBin}'. Bin does not exists in Warehouse '{ToWhse}'.");
            }
            else if (!String.IsNullOrWhiteSpace(ToBin))
                throw new Exception($"Line [{WMSTransId}]. Invalid 'To' Bin Code '{ToBin}'. Warehouse '{ToWhse}' not managed by bin.");


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

        public int GetBaseEntry(SAPbobsCOM.Company company, string headerWMSTransId)
        {
            if (_baseWtqEntry != null) return _baseWtqEntry ?? 0;

            IsBaseEntryExists(company, headerWMSTransId, out int baseEntry, out _, out _);
            return baseEntry;
        }

        public int GetBaseLine(SAPbobsCOM.Company company, string headerWMSTransId)
        {
            if (_baseWtqLine != null) return _baseWtqLine ?? 0;

            IsBaseEntryExists(company, headerWMSTransId, out _, out int baseLine, out _);
            return baseLine;
        }

        private bool IsBaseEntryExists(SAPbobsCOM.Company company, string headerWMSTransId, out int baseEntry, out int baseLine, out string result)
        {
            result = "";
            if(_baseWtqEntry!=null)
            {
                baseEntry = _baseWtqEntry??0;
                baseLine = _baseWtqLine??0;
                return _baseWtqEntry != 0;
            }

            baseEntry = 0;
            baseLine = -1;
            string sql = String.Format(Resource.MSSQL_Queries.OWTQ_GET_BASE_INFO_BY_WMS_ID, headerWMSTransId, RqWMSTransId);
            SAPbobsCOM.Recordset rs = null;
            try
            {
                rs = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                rs.DoQuery(sql);
                if (rs.RecordCount > 1) 
                    result = ($"Line [{WMSTransId}]. Multiple base line found for inventory request [{RqWMSTransId}]");
                else if(rs.RecordCount ==0)
                    result = ($"Line [{WMSTransId}]. No base line found for inventory request [{RqWMSTransId}]");

                if(result!="")
                {
                    _baseWtqEntry = 0;
                    _baseWtqLine = 0;
                    return false;
                }

                //Have exactly 1 match.
                _baseWtqEntry = (int)rs.Fields.Item("DocEntry").Value;
                _baseWtqLine = (int)rs.Fields.Item("LineNum").Value;
                baseEntry = _baseWtqEntry ?? 0;
                baseLine = _baseWtqLine ?? 0;
                return true;
            }
            finally
            {
                if (rs != null) SBOSupport.ReleaseComObject(rs);
            }

        }

        private bool IsWarehouseExists(SAPbobsCOM.Company company, string whse, out string result)
        {
            return SBOSupport.IsWarehouseExists(company, whse, out result);
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

        public int GetFromBinEntry(SAPbobsCOM.Company company)
        {
            if (_fromBinEntry != null)
                return _fromBinEntry ?? 0;

            if (String.IsNullOrWhiteSpace(FromBin))
            {
                _fromBinEntry = 0;
                return _fromBinEntry ?? 0;
            }

            _fromBinEntry = SBOSupport.GetBinEntry(company, FromBin, out string result);
            if (_fromBinEntry == 0) throw new Exception($"Line {WMSTransId}. {result}");
            return _fromBinEntry ?? 0;
        }

        public int GetToBinEntry(SAPbobsCOM.Company company)
        {
            if (_toBinEntry != null)
                return _toBinEntry ?? 0;

            if (String.IsNullOrWhiteSpace(ToBin))
            {
                _toBinEntry = 0;
                return _toBinEntry ?? 0;
            }

            _toBinEntry = SBOSupport.GetBinEntry(company, ToBin, out string result);
            if (_toBinEntry == 0) throw new Exception($"Line {WMSTransId}. {result}");
            return _toBinEntry ?? 0;
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

        public bool IsFromBinWarehouse(SAPbobsCOM.Company company)
        {
            if (_isFromBinWarehouse != null) return _isFromBinWarehouse ?? true;
            if (String.IsNullOrWhiteSpace(FromWhse))
            {
                _isFromBinWarehouse = false;
                return _isFromBinWarehouse ?? false;
            }

            _isFromBinWarehouse = SBOSupport.IsBinWarehouse(company, FromWhse);
            return _isFromBinWarehouse ?? true;
        }

        public bool IsToBinWarehouse(SAPbobsCOM.Company company)
        {
            if (_isToBinWarehouse != null) return _isToBinWarehouse ?? true;
            if (String.IsNullOrWhiteSpace(ToWhse))
            {
                _isToBinWarehouse = false;
                return _isToBinWarehouse ?? false;
            }

            _isToBinWarehouse = SBOSupport.IsBinWarehouse(company, ToWhse);
            return _isToBinWarehouse ?? true;
        }

        public bool IsValidSnb(SAPbobsCOM.Company company)
        {
            return SBOSupport.IsSnBExists(company, ItemCode, SNBCode, out string result);
        }

        public int GetSAPBaseType()
        {
            if (_baseWtqEntry != 0) return (int)SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest;
            else return -1;
        }
    }
}
