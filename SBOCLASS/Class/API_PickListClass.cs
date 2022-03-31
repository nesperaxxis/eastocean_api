using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SBOCLASS.Models.EOA;
using SAPbobsCOM;
namespace SBOCLASS.Class
{
    public class API_PickListClass
    {
        SAPbobsCOM.Company company { get; set; }
        SAPbobsCOM.PickLists HEADEROBJ { get; set; }
        string lastErrorMessage { get; set; }
        int lerrCode { get; set; }

        #region "SAPCONNECT"
        public bool connectToLoginCompany(string SQLServerName, string CompanyDB, string DBUserName, string DBPassword, string SBOUserName, string SBOPassword)
        {
            bool functionReturnValue = false;

            int lErrCode = 0;

            try
            {
                //// Initialize the Company Object.
                //// Create a new company object
                company = new SAPbobsCOM.Company();

                //// Set the mandatory properties for the connection to the database.
                //// To use a remote Db Server enter his name instead of the string "(local)"
                //// This string is used to work on a DB installed on your local machine

                company.Server = SQLServerName;
                company.CompanyDB = CompanyDB;
                company.UserName = SBOUserName;
                company.Password = SBOPassword;
                company.language = SAPbobsCOM.BoSuppLangs.ln_English;

                //// Use Windows authentication for database server.
                //// True for NT server authentication,
                //// False for database server authentication.
                company.UseTrusted = false;

                if (ConstantClass.SQLVersion == 2008)
                {
                    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                }
                else if (ConstantClass.SQLVersion == 2012)
                {
                    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                }
                else if (ConstantClass.SQLVersion == 2014)
                {
                    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                }
                else if (ConstantClass.SQLVersion == 2016)
                {
                    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                }
                else if (ConstantClass.SQLVersion == 2017)
                {
                    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
                }
                else if (ConstantClass.SQLVersion == 2019)
                {
                    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                }

                company.DbUserName = DBUserName;
                company.DbPassword = DBPassword;

                //// connect
                lErrCode = company.Connect();

                //// Check for errors during connect
                //sapCompany.GetLastError(lErrCode,lastErrorMessage_Out);
                if (lErrCode != 0)
                {
                    lastErrorMessage = "SAP Connection Error : " + company.GetLastErrorDescription();
                    functionReturnValue = false;
                }
                else
                {
                    functionReturnValue = true;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return functionReturnValue;
        }
        #endregion


        public PostObjectResult POSTPickList(PickList pickList, string SQLConnection)
        {
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from \"OPKL\" where Canceled='N' and  isnull(\"U_AXC_EXTID\",'') = '{ pickList.WMSTransId}'", company));
                    pickList.ValidateLine(company);

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{pickList.WMSTransId}' Already Exists in SAP Business One! [OPKL]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {

                        HEADEROBJ = CreateUnPickedPickList_Ex(pickList);
                        string xml = HEADEROBJ.GetAsXML();
                        company.StartTransaction();
                        int ierr = HEADEROBJ.Add();
                        if (ierr != 0) throw new Exception($"Create PickList failed. {company.GetLastErrorDescription()}");
                        int pickListEntry = int.Parse(company.GetNewObjectKey());
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oPickLists) as SAPbobsCOM.PickLists;
                        HEADEROBJ.GetByKey(pickListEntry);
                        string pickListString = HEADEROBJ.GetAsXML();
                        //Lines -> 1 line = multiple Batch = multiple Bin
                        //Group the lines by BaseKey and BaseLine first.
                        //For each unique line, if the Item is managed by Serial/Batch -> Further group by SnBCode
                        //For each unique SnbCode if the warehouse is managed by Bin -> group it by BinCode
                        //Sum the quantity as Picked Quantity.

                        var uniqueLines = pickList.Lines                            
                            .GroupBy(x => new { x.BaseKey, x.BaseLine, BaseType = x.GetSAPBaseType(), 
                                x.Whse, IsBinWarehouse = x.IsBinWarehouse(company), 
                                x.ItemCode, IsItemManagedBySnB = x.IsItemManagedBySnB(company),x.PickedQty}).Select(x => x);
                      
                        
                        int intLineCount = 0;
                        var pickLine = HEADEROBJ.Lines;
                        bool isSnBInvolved = false;
                        foreach (var line in uniqueLines)
                        {
                            if(line.Key.PickedQty > 0)
                            {
                                //Locate this unique baseline and key in existing picklist.
                                bool found = false;
                                for (int i = 0; i < pickLine.Count; i++)
                                {
                                    pickLine.SetCurrentLine(i);
                                    if (pickLine.BaseObjectType == line.Key.BaseType.ToString() &&
                                        pickLine.OrderEntry == line.Key.BaseKey &&
                                        pickLine.OrderRowID == line.Key.BaseLine)
                                    {

                                        found = true;
                                        break;
                                    }
                                }


                                if (!found)
                                    throw new Exception($"Line [{line.Select(x => x.WMSTransId).FirstOrDefault()}]. Invalid base line. {line.Key.BaseType} - {line.Key.BaseKey} - {line.Key.BaseLine}");

                                pickLine.UserFields.Fields.Item("U_AXC_EXTID").Value = String.Join(" ; ", line.Select(x => x.WMSTransId));
                                pickLine.PickedQuantity = line.Sum(x => x.PickedQty);
                                int snbLineNum = -1;
                                int binLineNum = -1;
                                var binLine = pickLine.BinAllocations;

                                if (line.Key.IsItemManagedBySnB == "S")
                                {
                                    //Group this line by SnbCode
                                    isSnBInvolved = true;
                                    var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                    var pickLineSerial = pickLine.SerialNumbers;
                                    foreach (var snbLine in thisLineSnBCodes)
                                    {
                                        snbLineNum++;
                                        if (snbLineNum > 0)
                                            pickLineSerial.Add();

                                        pickLineSerial.BaseLineNumber = intLineCount;
                                        pickLineSerial.InternalSerialNumber = snbLine.Key;
                                        pickLineSerial.Quantity = snbLine.Sum(x => x.PickedQty);

                                        if (pickLineSerial.Quantity > 1) throw new Exception($"Item '{line.Key.ItemCode}'. Serial Number '{snbLine.Key}'. Total quantity must be 1.");

                                        if (line.Key.IsBinWarehouse)
                                        {
                                            //Get the bins for this serial numbers
                                            var snbBinLines = snbLine.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                                            foreach (var snbbinLine in snbBinLines)
                                            {
                                                binLineNum++;
                                                if (binLineNum > 0) binLine.Add();
                                                binLine.BaseLineNumber = intLineCount;
                                                binLine.SerialAndBatchNumbersBaseLine = snbLineNum;
                                                binLine.BinAbsEntry = snbbinLine.Key;
                                                binLine.Quantity = snbbinLine.Sum(x => x.PickedQty);
                                            }
                                        }
                                    }
                                }
                                else if (line.Key.IsItemManagedBySnB == "B")
                                {
                                    isSnBInvolved = true;
                                    //managed by batch
                                    //Group this line by SnbCode
                                    var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                    var pickLineBatch = pickLine.BatchNumbers;
                                    foreach (var snbLine in thisLineSnBCodes)
                                    {
                                        snbLineNum++;
                                        if (snbLineNum > 0)
                                            pickLineBatch.Add();

                                        pickLineBatch.BaseLineNumber = intLineCount;
                                        pickLineBatch.BatchNumber = snbLine.Key;
                                        //pickLineBatch.InternalSerialNumber = snbLine.Key;
                                        pickLineBatch.Quantity = snbLine.Sum(x => x.PickedQty);

                                        if (line.Key.IsBinWarehouse)
                                        {
                                            //Get the bins for this serial numbers
                                            var snbBinLines = snbLine.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                                            foreach (var snbbinLine in snbBinLines)
                                            {
                                                binLineNum++;
                                                if (binLineNum > 0) binLine.Add();
                                                binLine.BaseLineNumber = intLineCount;
                                                binLine.SerialAndBatchNumbersBaseLine = snbLineNum;
                                                binLine.BinAbsEntry = snbbinLine.Key;
                                                binLine.Quantity = snbbinLine.Sum(x => x.PickedQty);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //Non batch
                                    //assign the bin only
                                    if (line.Key.IsBinWarehouse)
                                    {
                                        //Get the bins for this serial numbers
                                        var binLines = line.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                                        foreach (var nonSnBbinLine in binLines)
                                        {
                                            binLineNum++;
                                            if (binLineNum > 0) binLine.Add();
                                            binLine.BaseLineNumber = intLineCount;
                                            binLine.BinAbsEntry = nonSnBbinLine.Key;
                                            binLine.Quantity = nonSnBbinLine.Sum(x => x.PickedQty);
                                        }
                                    }
                                }

                                intLineCount += 1;
                            }
                            
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();
                        //Before we add the picklist, if SNB is involved, we need to allocate the SNB to the base Sales Order document
                        if (isSnBInvolved )
                            UpdateBaseSalesOrder(company, pickList);

                        lerrCode = HEADEROBJ.Update();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                //lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from \"OPKL\" where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_PICK_LIST, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw; 
#endif
                            }
                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + pickList.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
;                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                    }
                    catch
                    { }

                    AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_PICK_LIST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), pickList.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(pickList), response.Remark, response.Status);

                    if (company.InTransaction)
                        company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {lastErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting PickList failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_PICK_LIST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), pickList.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(pickList), response.Remark, response.Status);

            }
            finally
            {
                if (company == null && company.InTransaction) company.EndTransaction(BoWfTransOpt.wf_RollBack);
                try
                { System.Runtime.InteropServices.Marshal.ReleaseComObject(company); }
                catch
                { }
            }

            return response;
        }

        private SAPbobsCOM.PickLists CreateUnPickedPickList(SAPbobsCOM.BoObjectTypes baseType, List<int> docEntries, int user, DateTime pickDate, string remark, string extId)
        {
            SAPbobsCOM.PickLists pickList = company.GetBusinessObject(BoObjectTypes.oPickLists) as SAPbobsCOM.PickLists;
            if (user != 0) pickList.OwnerCode = user;
            pickList.PickDate = pickDate;
            pickList.Remarks = remark;
            pickList.UserFields.Fields.Item("U_AXC_EXTID").Value = extId;

            //Get the base document information.
            string tableName = SBOSupport.GetTableName(baseType);
            string sql = string.Format(Resource.MSSQL_Queries.OPKL_GET_BASE_LINES, tableName, string.Join(",", docEntries));
            SAPbobsCOM.Recordset rs = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            rs.DoQuery(sql);

            for(int i=0; i<rs.RecordCount; i++)
            {
                if (i > 0) pickList.Lines.Add();
                pickList.Lines.BaseObjectType = rs.Fields.Item("ObjType").Value.ToString();
                pickList.Lines.OrderEntry = (int)rs.Fields.Item("DocEntry").Value;
                pickList.Lines.OrderRowID = (int)rs.Fields.Item("LineNum").Value;
                pickList.Lines.ReleasedQuantity = (Double)rs.Fields.Item("OpenQty").Value;

                rs.MoveNext();
            }

            return pickList;

        }

        private SAPbobsCOM.PickLists CreateUnPickedPickList_Ex(PickList pickListData)
        {
            SAPbobsCOM.PickLists pickList = company.GetBusinessObject(BoObjectTypes.oPickLists) as SAPbobsCOM.PickLists;
            if (pickListData.User != 0) pickList.OwnerCode = pickListData.User;
            pickList.PickDate = pickListData.GetPickDate();
            pickList.Remarks = pickList.Remarks;
            pickList.UserFields.Fields.Item("U_AXC_EXTID").Value = pickListData.WMSTransId;

            //Create all the lines, from all the base entry.Full OpenQUantity
            var uniqueBaseEntries = pickListData.Lines.
                GroupBy(x => new {
                    x.BaseKey,
                    BaseType = x.GetSAPBaseType()
                }).Select(x => x);

            int iLineCount = 0;
            var binLine = pickList.Lines.BinAllocations;
            foreach (var baseEntry in uniqueBaseEntries)
            {
                //Get the base document information.
                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)baseEntry.Key.BaseType);
                string sql = string.Format(Resource.MSSQL_Queries.OPKL_GET_BASE_LINES, tableName, baseEntry.Key.BaseKey);
                SAPbobsCOM.Recordset rs = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                rs.DoQuery(sql);

                for (int i = 0; i < rs.RecordCount; i++)
                {
                    if (iLineCount > 0) pickList.Lines.Add();
                    pickList.Lines.BaseObjectType = rs.Fields.Item("ObjType").Value.ToString();
                    pickList.Lines.OrderEntry = (int)rs.Fields.Item("DocEntry").Value;
                    pickList.Lines.OrderRowID = (int)rs.Fields.Item("LineNum").Value;
                    pickList.Lines.ReleasedQuantity = (Double)rs.Fields.Item("OpenQty").Value;

                    iLineCount++;
                    rs.MoveNext();
                }
            }

            return pickList;

        }

        private SAPbobsCOM.PickLists CreateUnPickedPickList(PickList pickListData)
        {
            SAPbobsCOM.PickLists pickList = company.GetBusinessObject(BoObjectTypes.oPickLists) as SAPbobsCOM.PickLists;
            if (pickListData.User != 0) pickList.OwnerCode = pickListData.User;
            pickList.PickDate = pickListData.GetPickDate();
            pickList.Remarks = pickList.Remarks;
            pickList.UserFields.Fields.Item("U_AXC_EXTID").Value = pickListData.WMSTransId;

            //Create the lines, group by BaseType, BaseEntry, LineNum
            //If the whs is bin and item is managed by SnB, can only release the qty that is already picked, because to release bin + batch, the bin and qty must alreadt allocated.

            //Group the lines by BaseType, BaseKey and BaseLine first.
            //For each unique line, if the Item is managed by Serial/Batch -> Further group by SnBCode
            //For each unique SnbCode if the warehouse is managed by Bin -> group it by BinCode
            //Sum the quantity as Picked Quantity.
            var uniqueBaseEntries = pickListData.Lines.
                GroupBy(x => new {
                    x.BaseKey,
                    BaseType = x.GetSAPBaseType(),
                }).Select(x => x);

            int iLineCount = 0;
            int binLineNum = -1;
            var binLine = pickList.Lines.BinAllocations;
            foreach (var baseEntry in uniqueBaseEntries)
            {
                //if base document does not contain Bin and SnB, copy the whole base document.
                var hasBinAndSnb = baseEntry.Where(x => x.IsBinWarehouse(company) == true && x.IsItemManagedBySnB(company) != "").Count()>0;
                if(!hasBinAndSnb)
                {
                    //Get the base document information.
                    string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)baseEntry.Key.BaseType);
                    string sql = string.Format(Resource.MSSQL_Queries.OPKL_GET_BASE_LINES, tableName, baseEntry.Key.BaseKey);
                    SAPbobsCOM.Recordset rs = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                    rs.DoQuery(sql);

                    for (int i = 0; i < rs.RecordCount; i++)
                    {
                        if (iLineCount > 0) pickList.Lines.Add();
                        pickList.Lines.BaseObjectType = rs.Fields.Item("ObjType").Value.ToString();
                        pickList.Lines.OrderEntry = (int)rs.Fields.Item("DocEntry").Value;
                        pickList.Lines.OrderRowID = (int)rs.Fields.Item("LineNum").Value;
                        pickList.Lines.ReleasedQuantity = (Double)rs.Fields.Item("OpenQty").Value;

                        iLineCount++;
                        rs.MoveNext();
                    }
                } else
                {
                    //Has SNB and Bin - Can only release whatever is picked.
                    var uniqueLines = baseEntry.
                        GroupBy(x => new {
                            x.BaseKey,
                            x.BaseLine,
                            BaseType = x.GetSAPBaseType(),
                            x.Whse,
                            IsBinWarehouse = x.IsBinWarehouse(company),
                            x.ItemCode,
                            IsItemManagedBySnB = x.IsItemManagedBySnB(company),
                            IsSalesUOM = x.IsSalesUOM(company, out _) || x.GetItemNumInSale(company) == 1.0,
                            NumInSale = x.GetItemNumInSale(company)
                        }).Select(x => x);

                    binLineNum = -1;
                    int snbLineNum = -1;
                    foreach (var line in uniqueLines)
                    {
                        snbLineNum = -1;
                        binLineNum = -1;

                        if (iLineCount > 0) pickList.Lines.Add();
                        pickList.Lines.BaseObjectType = line.Key.BaseType.ToString();
                        pickList.Lines.OrderEntry = line.Key.BaseKey;
                        pickList.Lines.OrderRowID = line.Key.BaseLine;
                        pickList.Lines.ReleasedQuantity = line.Sum(x => x.PickedQty);

                        //Assigned the batch and bin if the item SNB and Warehouse is Bin

                        //if (line.Key.IsItemManagedBySnB == "S" && line.Key.IsBinWarehouse)
                        //{
                        //    //Group this line by SnbCode
                        //    var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                        //    var docLineSerial = pickList.Lines.SerialNumbers;
                        //    foreach (var snbLine in thisLineSnBCodes)
                        //    {
                        //        snbLineNum++;
                        //        if (snbLineNum > 0)
                        //            docLineSerial.Add();

                        //        docLineSerial.BaseLineNumber = iLineCount;
                        //        docLineSerial.BatchID = snbLine.Key;
                        //        docLineSerial.Quantity = snbLine.Sum(x => x.PickedQty);

                        //        if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.Key.ItemCode}'. Serial Number '{snbLine.Key}'. Total quantity must be 1.");

                        //        //Get the bins for this serial numbers
                        //        var snbBinLines = snbLine.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                        //        foreach (var snbbinLine in snbBinLines)
                        //        {
                        //            binLineNum++;
                        //            if (binLineNum > 0) binLine.Add();
                        //            binLine.BaseLineNumber = iLineCount;
                        //            binLine.SerialAndBatchNumbersBaseLine = snbLineNum;
                        //            binLine.BinAbsEntry = snbbinLine.Key;
                        //            binLine.Quantity = snbbinLine.Sum(x => x.PickedQty);
                        //        }
                        //    }
                        //}
                        //else if (line.Key.IsItemManagedBySnB == "B" && line.Key.IsBinWarehouse)
                        //{
                        //    //managed by batch 
                        //    //Group this line by SnbCode
                        //    var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                        //    var docLineBatch = pickList.Lines.BatchNumbers;
                        //    foreach (var snbLine in thisLineSnBCodes)
                        //    {
                        //        snbLineNum++;
                        //        if (snbLineNum > 0)
                        //            docLineBatch.Add();

                        //        docLineBatch.BaseLineNumber = iLineCount;
                        //        docLineBatch.BatchNumber = snbLine.Key;
                        //        docLineBatch.Quantity = snbLine.Sum(x => x.PickedQty);

                        //        //Get the bins for this serial numbers
                        //        var snbBinLines = snbLine.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                        //        foreach (var snbbinLine in snbBinLines)
                        //        {
                        //            binLineNum++;
                        //            if (binLineNum > 0) binLine.Add();
                        //            binLine.BaseLineNumber = iLineCount;
                        //            binLine.SerialAndBatchNumbersBaseLine = snbLineNum;
                        //            binLine.BinAbsEntry = snbbinLine.Key;
                        //            binLine.Quantity = snbbinLine.Sum(x => x.PickedQty);
                        //        }
                        //    }
                        //}


                        iLineCount++;
                    }
                }
            }

            return pickList;

        }

        private void UpdateBaseSalesOrder(SAPbobsCOM.Company company, PickList pickListData)
        {
            //Update the Base Sales Order and allocate the SnB.
            //Only need to update if it involve snb
            var snbPickLines = pickListData.Lines.Where(x => x.IsItemManagedBySnB(company) != "").Select(x => x);
            //group the pick list by base document 
            var baseDocuments = snbPickLines.GroupBy(x => new { BaseType = x.GetSAPBaseType(), x.BaseKey }).Select(x=>x);
            dynamic sapDocLines;
            SAPbobsCOM.BatchNumbers sapDocLineBatches;
            SAPbobsCOM.SerialNumbers sapDocLineSerials;
            foreach(var baseDoc in baseDocuments)
            {
                SAPbobsCOM.Documents doc = company.GetBusinessObject((BoObjectTypes)baseDoc.Key.BaseType) as SAPbobsCOM.Documents;
                if (!doc.GetByKey(baseDoc.Key.BaseKey)) continue;
                bool isUpdated = false;

                sapDocLines = doc.Lines;
                for (int i = 0; i < sapDocLines.Count; i++)
                {
                    sapDocLines.SetCurrentLine(i);

                    List<PickListDetail> lines = null;
                    switch (baseDoc.Key.BaseType)
                    {
                        case 17:
                        case 13:
                        case 1250000001:
                            lines = snbPickLines.Where(x => x.BaseLine == sapDocLines.LineNum).ToList();
                            break;
                        default:
                            lines = snbPickLines.Where(x => x.BaseLine == sapDocLines.LineNumber).ToList();
                            break;
                    }

                    if (lines == null)
                        continue;

                    List<String> appliedNumbers = new List<string>();
                    if ((lines.FirstOrDefault()?.IsItemManagedBySnB(company)??"") == "B")
                    {
                        sapDocLineBatches = doc.Lines.BatchNumbers;
                        for (int batchLineNum = 0; batchLineNum < sapDocLineBatches.Count; batchLineNum++)
                        {
                            sapDocLineBatches.SetCurrentLine(batchLineNum);
                            String distNumber = sapDocLineBatches.BatchNumber;
                            if (appliedNumbers.Contains(distNumber))
                            {
                                sapDocLineBatches.Quantity = 0;
                                isUpdated = true;
                                continue;
                            }

                            var pickedBatch = lines.Where(x => x.SNBCode == distNumber);
                            if (pickedBatch != null)
                            {
                                var pickedQty = pickedBatch.Sum(x => x.PickedQty);
                                if (sapDocLineBatches.Quantity < pickedQty)
                                {
                                    sapDocLineBatches.Quantity = pickedQty;
                                    isUpdated = true;
                                }
                                if (!appliedNumbers.Contains(distNumber))
                                    appliedNumbers.Add(distNumber);
                            }
                            else
                            {
                                //if (batches.Quantity != 0)
                                //{
                                //    batches.Quantity = 0;
                                //    isUpdated = true;
                                //}
                            }
                        }

                        //Add non existent batch
                        foreach (var pickedBatch in lines)
                        {
                            if (appliedNumbers.Contains(pickedBatch.SNBCode))
                                continue;

                            sapDocLineBatches.SetCurrentLine(sapDocLineBatches.Count - 1);
                            if (!String.IsNullOrEmpty(sapDocLineBatches.BatchNumber))
                                sapDocLineBatches.Add();

                            sapDocLineBatches.BatchNumber = pickedBatch.SNBCode;
                            sapDocLineBatches.Quantity = pickedBatch.PickedQty;
                            isUpdated = true;
                        }

                    }
                    else if ((lines.FirstOrDefault()?.IsItemManagedBySnB(company)??"" )== "S")
                    {
                        sapDocLineSerials = doc.Lines.SerialNumbers;
                        for (int serialLineNum = 0; serialLineNum < sapDocLineSerials.Count; serialLineNum++)
                        {
                            sapDocLineSerials.SetCurrentLine(serialLineNum);
                            String distNumber = sapDocLineSerials.InternalSerialNumber;
                            if (appliedNumbers.Contains(distNumber))
                            {
                                sapDocLineSerials.Quantity = 0;
                                isUpdated = true;
                                continue;
                            }
                            var pickedSerial = lines.Where(x => x.SNBCode == distNumber);
                            if (pickedSerial != null)
                            {
                                var pickedQty = pickedSerial.Sum(x => x.PickedQty);
                                if (sapDocLineSerials.Quantity != pickedQty)
                                {
                                    sapDocLineSerials.Quantity = pickedQty;
                                    isUpdated = true;
                                }
                                if (!appliedNumbers.Contains(distNumber))
                                    appliedNumbers.Add(distNumber);
                            }
                            else
                            {
                                //if (serials.Quantity != 0)
                                //{
                                //    serials.Quantity = 0;
                                //    isUpdated = true;
                                //}
                            }

                        }

                        //Add non existent serials
                        foreach (PickListDetail pickedSerials in lines)
                        {
                            if (appliedNumbers.Contains(pickedSerials.SNBCode))
                                continue;

                            sapDocLineSerials.SetCurrentLine(sapDocLineSerials.Count - 1);
                            if (sapDocLineSerials.SystemSerialNumber != 0)
                                sapDocLineSerials.Add();

                            sapDocLineSerials.SystemSerialNumber = SBOSupport.GetSerialSysNumber(company, pickedSerials.ItemCode, pickedSerials.SNBCode);
                            sapDocLineSerials.Quantity = pickedSerials.PickedQty;
                            isUpdated = true;
                        }

                    }
                }

                if (isUpdated)
                {
                    int ierr = doc.Update();
                    if (ierr != 0)
                    {
                        throw new Exception (company.GetLastErrorDescription());
                    }
                }
            }
        }



    }
}
