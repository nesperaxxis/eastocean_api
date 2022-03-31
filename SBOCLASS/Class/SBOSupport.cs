using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using SAPbobsCOM;
namespace SBOCLASS.Class
{
    public  static class SBOSupport
    {
        public enum Operation { POST, PUT, GET, PATCH, DELETE }
        public static string GETSINGLEVALUE(string StrQuery,SAPbobsCOM.Company SAPCompany)
        {
            SAPbobsCOM.Recordset oRecSet = null;
            try
            {
                SAPbobsCOM.Company company = SAPCompany;
                oRecSet = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                oRecSet.DoQuery(StrQuery);
                return Convert.ToString(oRecSet.Fields.Item(0).Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oRecSet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet);
            }
        }
        public static string GETSINGLEVALUESQL(string StrQuery,string ConnectionString)
        {
            string retValue = string.Empty;
            try
            {
                SQLClass sqlCls = new SQLClass();
                sqlCls.ConnectionString = ConnectionString;
                retValue = Convert.ToString(sqlCls.ExecuteScalar(StrQuery));
                return retValue;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool IsCardCodeExists(SAPbobsCOM.Company company, string cardCode, out string result)
        {
            bool exists;
            string sql = $"SELECT ISNULL(MAX(\"CardCode\"),'') FROM OCRD WHERE \"CardCode\" = '{cardCode.Replace("'", "''")}'";
            string sapCardCode = GETSINGLEVALUE(sql, company);
            if (String.IsNullOrWhiteSpace(sapCardCode))
            {
                exists = false;
                result = $"Business Partner '{cardCode}' does not exists. [OCRD.CardCode].";
            }
            else 
            {
                exists = true;
                result = sapCardCode;
            }

            return exists;
        }

        public static bool IsSlpCodeExists(SAPbobsCOM.Company company, string slpCode, out string result)
        {
            bool exists;
            string sql = $"SELECT ISNULL(MAX(\"SlpdCode\"),'') FROM OSLP WHERE \"SlpName\" = '{slpCode.Replace("'", "''")}'";
            string sapSlpCode = GETSINGLEVALUE(sql, company);
            if (String.IsNullOrWhiteSpace(sapSlpCode))
            {
                exists = false;
                result = $"Business Partner '{slpCode}' does not exists. [OCRD.CardCode].";
            }
            else
            {
                exists = true;
                result = sapSlpCode;
            }

            return exists;
        }

        public static bool IsWarehouseExists(SAPbobsCOM.Company company, string whsCode, out string result)
        {
            bool exists = true;
            result = "";
            string sql = $"SELECT ISNULL(MAX(\"Locked\"),'') FROM OWHS WHERE \"WhsCode\" = '{whsCode.Replace("'", "''")}'";
            string whsStatus = GETSINGLEVALUE(sql, company);
            if (String.IsNullOrWhiteSpace(whsStatus))
            {
                exists = false;
                result = $"Warehouse '{whsCode}' does not exists. [OWHS.WhsCode].";
            }
            else if (whsStatus == "Y")
            {
                exists = false;
                result = $"Warehouse '{whsCode}' is locked for transaction. [OWHS.WhsCode]";
            }

            return exists;
        }

        public static int GetBinEntry(SAPbobsCOM.Company company, string binCode, out string result)
        {
            result = "";
            if (String.IsNullOrWhiteSpace(binCode))
                return 0;

            string sql = $"SELECT ISNULL(MAX(\"AbsEntry\"),0) FROM OBIN WHERE \"BinCode\" = '{binCode.Replace("'", "'''")}' ";
            int binEntry = Convert.ToInt32(SBOSupport.GETSINGLEVALUE(sql, company));

            if (binEntry == 0)  result = ($"Bin Code '{binCode}' does not exists.");
            return binEntry;
        }

        public static bool IsItemExists(SAPbobsCOM.Company company, string itemCode)
        {
            bool isExists = true;
            if (String.IsNullOrWhiteSpace(itemCode))
                return false;
            string sql = $"SELECT COUNT('') FROM OITM WHERE \"ItemCode\" = '{itemCode.Replace("'", "''")}' ";
            isExists = Convert.ToInt32(SBOSupport.GETSINGLEVALUE(sql, company)) > 0;
            return isExists;
        }

        public static string IsItemManagedBySnB(SAPbobsCOM.Company company, string itemCode)
        {
            string sql = $"SELECT ISNULL(MAX(CASE WHEN \"ManSerNum\" = 'Y' THEN 'S' WHEN \"ManBtchNum\" = 'Y' THEN 'B' ELSE '' END),'') FROM OITM WHERE \"ItemCode\" = '{itemCode.Replace("'", "''")}' ";
            return GETSINGLEVALUE(sql, company).ToString();

        }

        public static bool IsBinWarehouse(SAPbobsCOM.Company company, string whsCode)
        {
            string sql = $"SELECT COUNT('') FROM OWHS WHERE BinActivat = 'Y' AND WhsCode = '{whsCode.Replace("'", "''")}' ";
            return Convert.ToInt32(SBOSupport.GETSINGLEVALUE(sql, company)) > 0;
        }

        public static bool IsSnBExists(SAPbobsCOM.Company company, string itemCode, string snbNumber, out string result)
        {
            result = "";
            string itemManageBy = IsItemManagedBySnB(company, itemCode);
            if (itemManageBy == "")
            {
                result = "Item not managed by Serial/Batch";
                return false;
            }

            string tableName = itemManageBy == "S" ? "OSR" : "OBT";
            string sql = $"SELECT ISNULL(MAX(T0.Quantity),0) FROM {tableName}Q T0 JOIN {tableName}N T1 ON T0.MdAbsEntry = T1.AbsEntry WHERE T0.ItemCode = '{itemCode.Replace("'","''")}' AND T1.DistNumber = '{snbNumber.Replace("'","''")}'";
            return Convert.ToInt32(SBOSupport.GETSINGLEVALUE(sql, company)) > 0;
        }

        public static double GetSnBOnHandQty(SAPbobsCOM.Company company, string itemCode, string snbNumber, string whsCode, string binCode, out string result)
        {
            result = "";
            string itemManageBy = IsItemManagedBySnB(company, itemCode);
            if (itemManageBy == "")
            {
                result = "Item not managed by Serial/Batch";
                return 0;
            }

            string sql;
            if (!String.IsNullOrWhiteSpace(binCode))
            {
                sql = itemManageBy == "S" ? Resource.MSSQL_Queries.OSBQ_GET_ONHAND_QTY : Resource.MSSQL_Queries.OBBQ_GET_ONHAND_QTY;
                sql = String.Format(sql, itemCode.Replace("'", "''"), snbNumber.Replace("'", "''"), binCode.Replace("'", "''"));
            } else
            {
                //Non bin
                sql = itemManageBy == "S" ? Resource.MSSQL_Queries.OSRQ_GET_ONHAND_QTY : Resource.MSSQL_Queries.OBTQ_GET_ONHAND_QTY;
                sql = String.Format(sql, itemCode.Replace("'", "''"), snbNumber.Replace("'", "''"), whsCode.Replace("'", "''"));

            }
            return Convert.ToInt32(SBOSupport.GETSINGLEVALUE(sql, company));
        }


        public static String GetItemInventoryUOM(SAPbobsCOM.Company company, string itemCode)
        {
            string sql = $"SELECT ISNULL(MAX(\"InvntryUom\"),'') FROM OITM WHERE ItemCode = '{itemCode.Replace("'","''")}'";
            return SBOSupport.GETSINGLEVALUE(sql, company).ToString();
        }

        public static String GetItemSalesUOM(SAPbobsCOM.Company company, string itemCode)
        {
            string sql = $"SELECT ISNULL(MAX(\"SalUnitMsr\"),'') FROM OITM WHERE ItemCode = '{itemCode.Replace("'", "''")}'";
            return SBOSupport.GETSINGLEVALUE(sql, company).ToString();
        }

        public static String GetItemPurchaseUOM(SAPbobsCOM.Company company, string itemCode)
        {
            string sql = $"SELECT ISNULL(MAX(\"BuyUnitMsr\"),'') FROM OITM WHERE ItemCode = '{itemCode.Replace("'", "''")}'";
            return SBOSupport.GETSINGLEVALUE(sql, company).ToString();
        }

        public static Double GetItemMasterDataNumInBuy(SAPbobsCOM.Company company, string itemCode)
        {
            string sql = $"SELECT ISNULL(MAX(\"NumInBuy\"),1) FROM OITM WHERE ItemCode = '{itemCode.Replace("'", "''")}'";
            return Convert.ToDouble(SBOSupport.GETSINGLEVALUE(sql, company));
        }

        public static Double GetItemNumInSale(SAPbobsCOM.Company company, string itemCode)
        {
            string sql = $"SELECT ISNULL(MAX(\"NumInSale\"),1) FROM OITM WHERE ItemCode = '{itemCode.Replace("'", "''")}'";
            return Convert.ToDouble(SBOSupport.GETSINGLEVALUE(sql, company));
        }

        public static Double GetItemNumPerMsr(SAPbobsCOM.Company company, int objType, int docEntry, int docLine, out string baseDocUom, out bool useBaseUnit)
        {
            string tableName = SBOSupport.GetTableName((BoObjectTypes)objType);
            string sql = $"SELECT NumPerMsr, UseBaseUn, unitMsr FROM {tableName}1 WHERE DocEntry = {docEntry} AND LineNum = {docLine}";
            var rs = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try
            {
                rs.DoQuery(sql);
                if (rs.RecordCount == 0)
                    throw new Exception($"Document line not found [{tableName}1]DocEntry = {docEntry}, LineNum = {docLine}.");

                baseDocUom = rs.Fields.Item("unitMsr").Value.ToString();
                useBaseUnit = rs.Fields.Item("UseBaseUn").Value.ToString().ToUpper() == "Y";
                return Convert.ToDouble(rs.Fields.Item("NumPerMsr").Value);
            }
            finally
            {
                SBOSupport.ReleaseComObject(rs);
            }
        }

        public static void SendAlert(SAPbobsCOM.Company company, bool isNew, string ObjName, String ObjType1, String Key1, String Code1, String ObjType2 = "", String Key2 = "", String Code2 = "")
        {

            SAPbobsCOM.Recordset oRS = null;
            SAPbobsCOM.CompanyService oCmpSrv = null;
            SAPbobsCOM.MessagesService oMessageService = null;
            SAPbobsCOM.Message oMessage = null;
            SAPbobsCOM.MessageDataColumns pMessageDataColumns = null;
            SAPbobsCOM.MessageDataColumn pMessageDataColumn = null;
            SAPbobsCOM.MessageDataLines oLines = null;
            SAPbobsCOM.MessageDataLine oLine = null;
            SAPbobsCOM.RecipientCollection oRecipientCollection = null;
            String Subject = "";
            String Text = "";
            String sSQL = "";
            String AlertFieldName = "";
            String[] ColumnNames = new String[0];
            String Status = "New";
            if (!isNew)
                Status = "Update";

            try
            {
                switch (ObjName)
                {
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_GRPO:
                        Subject = String.Format("{0} GRPO Received [{1}]", Status, Code1);
                        Text = String.Format("{0} GRPO has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALPDN";
                        ColumnNames = new String[] { "G.Receipt Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_PICK_LIST:
                        Subject = String.Format("{0} PickList [{1}]", Status, Code1);
                        Text = String.Format("{0} PickList has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALPKL";
                        ColumnNames = new String[] { "PickList Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_ISSUE_PROD:
                        Subject = String.Format("{0} Issue For Production [{1}]", Status, Code1);
                        Text = String.Format("{0} Issue for production has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALGEP";
                        ColumnNames = new String[] { "Prd Issue Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_RECPT_PROD:
                        Subject = String.Format("{0} Receipt From Production [{1}]", Status, Code1);
                        Text = String.Format("{0} Receipt from production has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALGNP";
                        ColumnNames = new String[] { "Prd Receipt Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_AR_RETURN:
                        Subject = String.Format("{0} Sales Returns [{1}]", Status, Code1);
                        Text = String.Format("{0} Sales return has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALRDN";
                        ColumnNames = new String[] { "Sales Return Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_DELIVERY_ORDER:
                        Subject = String.Format("{0} Delivery based on DO Retrn [{1}]", Status, Code1);
                        Text = String.Format("{0} Delivery Order has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALDLN";
                        ColumnNames = new String[] { "Delivery Number" };
                        break;

                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_TR_REQUEST:
                        Subject = String.Format("{0} Inventory Transfer Request [{1}]", Status, Code1);
                        Text = String.Format("{0} Inventory transfer request has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALWTQ";
                        ColumnNames = new String[] { "Trsfr Req Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_WHS_TRANSFER:
                        Subject = String.Format("{0} Inventory Transfer [{1}]", Status, Code1);
                        Text = String.Format("{0} Inventory transfer has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALWTR";
                        ColumnNames = new String[] { "Inv Trsfr Number" };
                        break;
                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_POS:
                        Subject = String.Format("{0} Stock Adj Increase [{1}]", Status, Code1);
                        Text = String.Format("{0} Stock Adjustment has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALIGN";
                        ColumnNames = new String[] { "G.Entry Number" };
                        break;

                    case Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_NEG:
                        Subject = String.Format("{0} Stock Adj Decrease [{1}]", Status, Code1);
                        Text = String.Format("{0} Stock Adjustment has been posted from web service", Status);
                        AlertFieldName = "U_AXC_ALIGE";
                        ColumnNames = new String[] { "G.Issue Number" };
                        break;

                    default:
                        throw new Exception(" Unsupported alert object. " + ObjName);
                }

                sSQL = $"SELECT T0.U_AXC_UNAME, T1.USER_CODE FROM \"@AXC_FTIS1\" T0 JOIN OUSR T1 ON T0.U_AXC_USRID = T1.USERID AND T1.GROUPS <> 99 WHERE T0.{AlertFieldName} = 'Y'";
                oRS = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                oRS.DoQuery(sSQL);
                if (oRS.RecordCount == 0)
                    return;             //No recipient.. no need to proceed.

                oCmpSrv = company.GetCompanyService();
                oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService) as SAPbobsCOM.MessagesService;
                oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage) as SAPbobsCOM.Message;

                if (Subject.Length > 254) Subject = Subject.Substring(0, 254);
                oMessage.Subject = Subject;
                oMessage.Text = Text;
                oRecipientCollection = oMessage.RecipientCollection;

                for (int iUser = 0; iUser < oRS.RecordCount; iUser++)
                {
                    oRecipientCollection.Add();
                    oRecipientCollection.Item(iUser).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    oRecipientCollection.Item(iUser).UserCode = oRS.Fields.Item("USER_CODE").Value.ToString();

                    oRS.MoveNext();
                }

                pMessageDataColumns = oMessage.MessageDataColumns;
                int iCol = 0;
                foreach (String columnName in ColumnNames)
                {
                    iCol++;
                    pMessageDataColumn = pMessageDataColumns.Add();
                    pMessageDataColumn.ColumnName = columnName;
                    pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES;
                    oLines = pMessageDataColumn.MessageDataLines;
                    oLine = oLines.Add();
                    if (iCol == 1)
                    {
                        oLine.Object = ObjType1;
                        oLine.ObjectKey = Key1;
                        oLine.Value = Code1;
                    }
                    else if (iCol == 2)
                    {
                        oLine.Object = ObjType2;
                        oLine.ObjectKey = Key2;
                        oLine.Value = Code2;
                    }
                }

                oMessageService.SendMessage(oMessage);

            }
            catch (Exception Ex)
            {
                AXC_OFTLG.GenerateLogRecord(company, ObjType1, Key1, Code1, "", Operation.POST, "", String.Format("Failed sending alert message {0}:{1}", ObjType1, Key1), false);

            }
            finally
            {
                ReleaseComObject(oRS);
                ReleaseComObject(oCmpSrv);
                ReleaseComObject(oMessageService);
                ReleaseComObject(oMessage);
                ReleaseComObject(pMessageDataColumns);
                ReleaseComObject(pMessageDataColumn);
                ReleaseComObject(oLines);
                ReleaseComObject(oLine);
                ReleaseComObject(oRecipientCollection);
            }

        }

        public static int GetSerialSysNumber(SAPbobsCOM.Company company, string itemCode, string distNumber)
        {
            String sql = String.Format("SELECT ISNULL(MAX(SysNumber),0) FROM OSRN WHERE ItemCode = '{0}' AND DistNumber = '{1}'", itemCode.Replace("'", "''"), distNumber.Replace("'", "''"));
            SAPbobsCOM.Recordset rs = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;

            return int.Parse(GETSINGLEVALUE(sql, company));
        }

        public static void ReleaseComObject(Object o)
        {
            if (o != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
        }

        public static string GetTableName(SAPbobsCOM.BoObjectTypes objType)
        {
            switch (objType)
            {
                case BoObjectTypes.oDrafts:
                    return "DRF";
                case BoObjectTypes.oPickLists:
                    return "PKL";
                case BoObjectTypes.oOrders:
                    return "RDR";
                case BoObjectTypes.oInvoices:
                    return "INV";
                case BoObjectTypes.oDeliveryNotes:
                    return "DLN";
                case BoObjectTypes.oPurchaseDeliveryNotes:
                    return "PDN";
                case BoObjectTypes.oReturns:
                    return "RDN";
                case BoObjectTypes.oPurchaseOrders:
                    return "POR";
                case BoObjectTypes.oStockTransfer:
                    return "WTR";
                case BoObjectTypes.oStockTransferDraft:
                    return "DRF";
                case BoObjectTypes.oInventoryTransferRequest:
                    return "WTQ";
                case BoObjectTypes.oInventoryGenEntry:
                    return "IGN";
                case BoObjectTypes.oInventoryGenExit:
                    return "IGE";
                case BoObjectTypes.oProductionOrders:
                    return "WOR";
                default:
                    throw new NotImplementedException("Object type not mapped.");
            }
        }

        public static double RoundSAPAmount(this double amount, SAPbobsCOM.Company company, string currency, SAPbobsCOM.RoundingContextEnum context)
        {
            SAPbobsCOM.DecimalData decimalData = company.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiDecimalData) as SAPbobsCOM.DecimalData;
            decimalData.Value = amount;
            decimalData.Context = context;
            if (!String.IsNullOrEmpty(currency))
                decimalData.Currency = currency;
            SAPbobsCOM.CompanyService oCS = company.GetCompanyService();
            SAPbobsCOM.RoundedData result = oCS.RoundDecimal(decimalData);
            return result.Value;
        }

    }

}
