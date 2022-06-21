using SAPbobsCOM;
using SBOCLASS.Models.EOA;
using System;
using System.Linq;
namespace SBOCLASS.Class
{
    public class API_DocumentClass
    {
        SAPbobsCOM.Company company { get; set; }
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
                //else if (ConstantClass.SQLVersion == 2017)
                //{
                //    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
                //}
                //else if (ConstantClass.SQLVersion == 2019)
                //{
                //    company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                //}

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
                    company.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;
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

        public PostObjectResult POSTGRPO(GRPO grpo, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "PDN";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} where Canceled='N' and  isnull(\"U_AXC_EXTID\",'') = '{ grpo.WMSTransId}'", company));
                    string cardCode = grpo.GetSAPCardCode(company);

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{grpo.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the GRPO lines first.
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes) as SAPbobsCOM.Documents;
                        //HEADEROBJ.GetByKey(444);
                        //string created = HEADEROBJ.GetAsXML();
                        grpo.ValidateLine(company);
                        //if (grpo.SlpCode != 0) HEADEROBJ.SalesPersonCode = grpo.SlpCode;
                        if (String.IsNullOrWhiteSpace(grpo.SlpCode)) HEADEROBJ.SalesPersonCode = Convert.ToInt32(grpo.GetSAPSlpCode(company));
                        HEADEROBJ.CardCode = cardCode;
                        HEADEROBJ.DocDate = grpo.GetPostDate();
                        HEADEROBJ.Comments = grpo.Remark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = grpo.WMSTransId;

                        //Lines -> 1 line = multiple Batch = multiple Bin
                        //Group the lines by BaseKey and BaseLine first.
                        //For each unique line, if the Item is managed by Serial/Batch -> Further group by SnBCode
                        //For each unique SnbCode if the warehouse is managed by Bin -> group it by BinCode
                        //Sum the quantity as Picked Quantity.

                        var uniqueLines = grpo.Lines.
                            GroupBy(x => new {
                                x.BaseKey,
                                x.BaseLine,
                                SAPBaseType = x.GetSAPBaseType(),
                                x.Whse,
                                IsBinWarehouse = x.IsBinWarehouse(company),
                                x.ItemCode,
                                IsItemManagedBySnB = x.IsItemManagedBySnB(company),
                                LineNumInBuy = x.GetItemNumInBuy(company),
                                IsPurchaseUOM = x.IsPurchaseUOM(company, out _) || x.GetItemNumInBuy(company)==1.0
                            }).Select(x => x);

                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in uniqueLines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();
                            if (line.Key.SAPBaseType != -1)
                            {
                                docLine.BaseType = line.Key.SAPBaseType;
                                docLine.BaseEntry = line.Key.BaseKey;
                                docLine.BaseLine = line.Key.BaseLine;
                            }
                            else
                            {
                                docLine.ItemCode = line.Key.ItemCode;
                                docLine.UseBaseUnits = line.Key.IsPurchaseUOM ? BoYesNoEnum.tNO : BoYesNoEnum.tYES;
                            }

                            docLine.WarehouseCode = line.Key.Whse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = String.Join(" ; ", line.Select(x => x.WMSTransId));
                            docLine.Quantity = line.Sum(x => x.Quantity);
                            int snbLineNum = -1;
                            int binLineNum = -1;
                            var binLine = docLine.BinAllocations;

                            if (line.Key.IsItemManagedBySnB == "S")
                            {
                                //Group this line by SnbCode
                                var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                var docLineSerial = docLine.SerialNumbers;
                                double totalQty = 0;
                                foreach (var snbLine in thisLineSnBCodes)
                                {
                                    snbLineNum++;
                                    if (snbLineNum > 0)
                                        docLineSerial.Add();

                                    docLineSerial.BaseLineNumber = intLineCount;
                                    docLineSerial.BatchID = snbLine.Key;
                                    docLineSerial.InternalSerialNumber = snbLine.Key;
                                    docLineSerial.ManufacturerSerialNumber = snbLine.Key;
                                    docLineSerial.Quantity = snbLine.Sum(x => x.Quantity); // (snbLine.Sum(x => x.Quantity) / line.Key.LineNumInBuy).RoundSAPAmount(company, "SGD", RoundingContextEnum.rcQuantity);
                                    totalQty += 1;
                                    if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.Key.ItemCode}'. Serial Number '{snbLine.Key}'. Total quantity must be 1.");

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
                                            binLine.Quantity = snbbinLine.Sum(x => x.Quantity);
                                        }
                                    }
                                }
                                docLine.Quantity = (totalQty / line.Key.LineNumInBuy).RoundSAPAmount(company, "SGD", RoundingContextEnum.rcQuantity);
                            }
                            else if (line.Key.IsItemManagedBySnB == "B")
                            {
                                //managed by batch
                                //Group this line by SnbCode
                                var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                var docLineBatch = docLine.BatchNumbers;
                                foreach (var snbLine in thisLineSnBCodes)
                                {
                                    snbLineNum++;
                                    if (snbLineNum > 0)
                                        docLineBatch.Add();

                                    docLineBatch.BaseLineNumber = intLineCount;
                                    docLineBatch.BatchNumber = snbLine.Key;
                                    docLineBatch.ManufacturerSerialNumber = snbLine.Key;
                                    docLineBatch.InternalSerialNumber = snbLine.Key;
                                    docLineBatch.Quantity = snbLine.Sum(x => (line.Key.IsPurchaseUOM?line.Key.LineNumInBuy:1.0) * x.Quantity).RoundSAPAmount(company, null, RoundingContextEnum.rcQuantity);

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
                                            binLine.Quantity = snbbinLine.Sum(x => (line.Key.IsPurchaseUOM ? line.Key.LineNumInBuy : 1.0) * x.Quantity).RoundSAPAmount(company, null, RoundingContextEnum.rcQuantity);
                                        }
                                    }
                                }
                            } else
                            {
                                //Not manage by serial or batch but have bin;
                                if (line.Key.IsBinWarehouse)
                                {
                                    //Get the bins for this row
                                    var binLines = line.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                                    foreach (var noBatchBinLine in binLines)
                                    {
                                        binLineNum++;
                                        if (binLineNum > 0) binLine.Add();
                                        binLine.BaseLineNumber = intLineCount;
                                        binLine.BinAbsEntry = noBatchBinLine.Key;
                                        binLine.Quantity = noBatchBinLine.Sum(x => (line.Key.IsPurchaseUOM ? line.Key.LineNumInBuy : 1.0) * x.Quantity).RoundSAPAmount(company, null, RoundingContextEnum.rcQuantity);
                                    }
                                }

                            }

                            intLineCount += 1;
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_GRPO, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_GRPO : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), grpo.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(grpo), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + grpo.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting GRPO failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_GRPO : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), grpo.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(grpo), response.Remark, response.Status);

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

        public PostObjectResult POSTIssueProd(IssueForProduction igeProd, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "IGE";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 JOIN {objectTableName}1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.Canceled='N' and T1.\"BaseType\" = 202 AND isnull(T0.\"U_AXC_EXTID\",'') = '{ igeProd.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{igeProd.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        igeProd.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oInventoryGenExit) as SAPbobsCOM.Documents;

                        HEADEROBJ.Reference2 = igeProd.Ref2;
                        HEADEROBJ.DocDate = igeProd.GetPostDate();
                        HEADEROBJ.Comments = igeProd.Remark;
                        if (!String.IsNullOrWhiteSpace(igeProd.JournalRemark)) HEADEROBJ.JournalMemo = igeProd.JournalRemark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = igeProd.WMSTransId;

                        //Lines -> 1 line = multiple Batch = multiple Bin
                        //Group the lines by BaseKey and BaseLine first.
                        //For each unique line, if the Item is managed by Serial/Batch -> Further group by SnBCode
                        //For each unique SnbCode if the warehouse is managed by Bin -> group it by BinCode
                        //Sum the quantity as Picked Quantity.

                        var uniqueLines = igeProd.Lines.
                            GroupBy(x => new {
                                x.BaseKey,
                                x.BaseLine,
                                SAPBaseType = x.GetSAPBaseType(),
                                x.Whse,
                                IsBinWarehouse = x.IsBinWarehouse(company),
                                x.ItemCode,
                                IsItemManagedBySnB = x.IsItemManagedBySnB(company),
                            }).Select(x => x);

                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in uniqueLines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            if (line.Key.SAPBaseType != -1)
                            {
                                docLine.BaseType = line.Key.SAPBaseType;
                                docLine.BaseEntry = line.Key.BaseKey;
                                if (line.Key.BaseLine > -1)
                                    docLine.BaseLine = line.Key.BaseLine;
                            }

                            //docLine.ItemCode = line.Key.ItemCode;     //Issue for production do not need to provide the item code
                            docLine.WarehouseCode = line.Key.Whse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = String.Join(" ; ", line.Select(x => x.WMSTransId));
                            docLine.Quantity = line.Sum(x => x.Quantity);

                            int snbLineNum = -1;
                            int binLineNum = -1;
                            var binLine = docLine.BinAllocations;

                            if (line.Key.IsItemManagedBySnB == "S")
                            {
                                //Group this line by SnbCode
                                var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                var docLineSerial = docLine.SerialNumbers;
                                foreach (var snbLine in thisLineSnBCodes)
                                {
                                    snbLineNum++;
                                    if (snbLineNum > 0)
                                        docLineSerial.Add();

                                    docLineSerial.BaseLineNumber = intLineCount;
                                    docLineSerial.BatchID = snbLine.Key;
                                    docLineSerial.InternalSerialNumber = snbLine.Key;
                                    docLineSerial.ManufacturerSerialNumber = snbLine.Key;
                                    docLineSerial.Quantity = snbLine.Sum(x => x.Quantity);

                                    if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.Key.ItemCode}'. Serial Number '{snbLine.Key}'. Total quantity must be 1.");

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
                                            binLine.Quantity = snbbinLine.Sum(x => x.Quantity);
                                        }
                                    }
                                }
                            }
                            else if (line.Key.IsItemManagedBySnB == "B")
                            {
                                //managed by batch
                                //Group this line by SnbCode
                                var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                var docLineBatch = docLine.BatchNumbers;
                                foreach (var snbLine in thisLineSnBCodes)
                                {
                                    snbLineNum++;
                                    if (snbLineNum > 0)
                                        docLineBatch.Add();

                                    docLineBatch.BaseLineNumber = intLineCount;
                                    docLineBatch.BatchNumber = snbLine.Key;
                                    docLineBatch.ManufacturerSerialNumber = snbLine.Key;
                                    docLineBatch.InternalSerialNumber = snbLine.Key;
                                    docLineBatch.Quantity = snbLine.Sum(x => x.Quantity);

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
                                            binLine.Quantity = snbbinLine.Sum(x => x.Quantity);
                                        }
                                    }
                                }
                            } else
                            {
                                //non SnB
                                if (line.Key.IsBinWarehouse)
                                {
                                    //Get the bins for this row
                                    var binLines = line.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                                    foreach (var noBatchBinLine in binLines)
                                    {
                                        binLineNum++;
                                        if (binLineNum > 0) binLine.Add();
                                        binLine.BaseLineNumber = intLineCount;
                                        binLine.BinAbsEntry = noBatchBinLine.Key;
                                        binLine.Quantity = noBatchBinLine.Sum(x =>  x.Quantity);
                                    }
                                }

                            }

                            intLineCount += 1;
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_ISSUE_PROD, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_ISSUE_PROD : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), igeProd.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(igeProd), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + igeProd.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Issue For Production failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_ISSUE_PROD : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), igeProd.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(igeProd), response.Remark, response.Status);

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

        public PostObjectResult POSTReceiptProd(ReceiptFromProduction ignProd, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "IGN";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 JOIN {objectTableName}1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.Canceled='N' and T1.\"BaseType\" = 202 AND isnull(T0.\"U_AXC_EXTID\",'') = '{ ignProd.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{ignProd.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        ignProd.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oInventoryGenEntry) as SAPbobsCOM.Documents;

                        HEADEROBJ.Reference2 = ignProd.Ref2;
                        HEADEROBJ.DocDate = ignProd.GetPostDate();
                        HEADEROBJ.Comments = ignProd.Remark;
                        if (!String.IsNullOrWhiteSpace(ignProd.JournalRemark)) HEADEROBJ.JournalMemo = ignProd.JournalRemark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = ignProd.WMSTransId;

                        //SAP Lines -> 1 line = multiple Batch = multiple Bin
                        //The data coming from WMS is 1 line = 1 batch and 1 bin.
                        //Group the lines by BaseKey and BaseLine first.
                        //For each unique line, if the Item is managed by Serial/Batch -> Further group by SnBCode
                        //For each unique SnbCode if the warehouse is managed by Bin -> group it by BinCode
                        //Sum the quantity as Picked Quantity.

                        var uniqueLines = ignProd.Lines.
                            GroupBy(x => new {
                                x.BaseKey,
                                x.BaseLine,
                                SAPBaseType = x.GetSAPBaseType(),
                                x.Whse,
                                IsBinWarehouse = x.IsBinWarehouse(company),
                                x.ItemCode,
                                IsItemManagedBySnB = x.IsItemManagedBySnB(company),
                                x.ReceiptType
                            }).Select(x => x);

                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in uniqueLines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            if (line.Key.SAPBaseType != -1)
                            {
                                docLine.BaseType = line.Key.SAPBaseType;
                                docLine.BaseEntry = line.Key.BaseKey;
                                if(line.Key.BaseLine>-1) docLine.BaseLine = line.Key.BaseLine;
                            }

                            //docLine.ItemCode = line.Key.ItemCode;         //ItemCode is not needed when receipt from production
                            docLine.WarehouseCode = line.Key.Whse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = String.Join(" ; ", line.Select(x => x.WMSTransId));
                            docLine.Quantity = line.Sum(x => x.Quantity);
                            if (line.Key.ReceiptType == "C")
                                docLine.TransactionType = BoTransactionTypeEnum.botrntComplete;
                            else if (line.Key.ReceiptType == "R")
                                docLine.TransactionType = BoTransactionTypeEnum.botrntReject;

                            int snbLineNum = -1;
                            int binLineNum = -1;
                            var binLine = docLine.BinAllocations;

                            if (line.Key.IsItemManagedBySnB == "S")
                            {
                                //Group this line by SnbCode
                                var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                var docLineSerial = docLine.SerialNumbers;
                                foreach (var snbLine in thisLineSnBCodes)
                                {
                                    snbLineNum++;
                                    if (snbLineNum > 0)
                                        docLineSerial.Add();

                                    docLineSerial.BaseLineNumber = intLineCount;
                                    docLineSerial.BatchID = snbLine.Key;
                                    docLineSerial.InternalSerialNumber = snbLine.Key;
                                    docLineSerial.ManufacturerSerialNumber = snbLine.Key;
                                    docLineSerial.Quantity = snbLine.Sum(x => x.Quantity);

                                    if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.Key.ItemCode}'. Serial Number '{snbLine.Key}'. Total quantity must be 1.");

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
                                            binLine.Quantity = snbbinLine.Sum(x => x.Quantity);
                                        }
                                    }
                                }
                            }
                            else if (line.Key.IsItemManagedBySnB == "B")
                            {
                                //managed by batch
                                //Group this line by SnbCode
                                var thisLineSnBCodes = line.GroupBy(x => x.SNBCode).Select(x => x);
                                var docLineBatch = docLine.BatchNumbers;
                                foreach (var snbLine in thisLineSnBCodes)
                                {
                                    snbLineNum++;
                                    if (snbLineNum > 0)
                                        docLineBatch.Add();

                                    docLineBatch.BaseLineNumber = intLineCount;
                                    docLineBatch.BatchNumber = snbLine.Key;
                                    docLineBatch.ManufacturerSerialNumber = snbLine.Key;
                                    docLineBatch.InternalSerialNumber = snbLine.Key;
                                    docLineBatch.Quantity = snbLine.Sum(x => x.Quantity);

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
                                            binLine.Quantity = snbbinLine.Sum(x => x.Quantity);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //non SnB
                                if (line.Key.IsBinWarehouse)
                                {
                                    //Get the bins for this row
                                    var binLines = line.GroupBy(x => x.GetBinEntry(company)).Select(x => x);
                                    foreach (var noBatchBinLine in binLines)
                                    {
                                        binLineNum++;
                                        if (binLineNum > 0) binLine.Add();
                                        binLine.BaseLineNumber = intLineCount;
                                        binLine.BinAbsEntry = noBatchBinLine.Key;
                                        binLine.Quantity = noBatchBinLine.Sum(x => x.Quantity);
                                    }
                                }

                            }

                            intLineCount += 1;
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_RECPT_PROD, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_RECPT_PROD : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), ignProd.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(ignProd), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + ignProd.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Receipt From Production failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_RECPT_PROD : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), ignProd.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(ignProd), response.Remark, response.Status);

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

        public PostObjectResult POSTStockAdjPos(StockAdjustmentPos ign, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "IGN";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 JOIN {objectTableName}1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.Canceled='N' and T1.\"BaseType\" = -1 AND isnull(T0.\"U_AXC_EXTID\",'') = '{ ign.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{ign.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        ign.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oInventoryGenEntry) as SAPbobsCOM.Documents;

                        HEADEROBJ.Reference2 = ign.Ref2;
                        HEADEROBJ.DocDate = ign.GetPostDate();
                        HEADEROBJ.Comments = ign.Remark;
                        if (!String.IsNullOrWhiteSpace(ign.JournalRemark)) HEADEROBJ.JournalMemo = ign.JournalRemark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = ign.WMSTransId;

                        //Create one for each line
                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in ign.Lines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            docLine.ItemCode = line.ItemCode;
                            if (!String.IsNullOrWhiteSpace(line.ItemName)) docLine.ItemDescription = line.ItemName;
                            docLine.WarehouseCode = line.Whse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = line.WMSTransId;
                            docLine.Quantity = line.Quantity;

                            if (line.IsItemManagedBySnB(company) == "S")
                            {
                                var docLineSerial = docLine.SerialNumbers;
                                docLineSerial.BaseLineNumber = intLineCount;
                                docLineSerial.InternalSerialNumber = line.SNBCode;
                                docLineSerial.Quantity = line.Quantity;
                                if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.ItemCode}'. Serial Number '{line.SNBCode}'. Total quantity must be 1.");
                            }
                            else if (line.IsItemManagedBySnB(company) == "B")
                            {
                                //managed by batch
                                var docLineBatch = docLine.BatchNumbers;
                                docLineBatch.BaseLineNumber = intLineCount;
                                docLineBatch.BatchNumber = line.SNBCode;
                                docLineBatch.Quantity = line.Quantity;
                            }

                            //Bin
                            if (line.IsBinWarehouse(company))
                            {
                                var docLineBin = docLine.BinAllocations;
                                docLineBin.BaseLineNumber = intLineCount;
                                docLineBin.BinAbsEntry = line.GetBinEntry(company);
                                docLineBin.Quantity = line.Quantity;
                                if (line.IsItemManagedBySnB(company) != "")
                                    docLineBin.SerialAndBatchNumbersBaseLine = 0;
                            }

                            intLineCount += 1;
                        }

                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_POS, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_POS : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), ign.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(ign), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + ign.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Stock Adjustment (+) failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_POS : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), ign.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(ign), response.Remark, response.Status);

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

        public PostObjectResult POSTStockAdjNeg(StockAdjustmentNeg ige, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "IGE";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 JOIN {objectTableName}1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.Canceled='N' and T1.\"BaseType\" = -1 AND isnull(T0.\"U_AXC_EXTID\",'') = '{ ige.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{ige.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        ige.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oInventoryGenExit) as SAPbobsCOM.Documents;

                        HEADEROBJ.Reference2 = ige.Ref2;
                        HEADEROBJ.DocDate = ige.GetPostDate();
                        HEADEROBJ.Comments = ige.Remark;
                        if (!String.IsNullOrWhiteSpace(ige.JournalRemark)) HEADEROBJ.JournalMemo = ige.JournalRemark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = ige.WMSTransId;

                        //Create one for each line
                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in ige.Lines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            docLine.ItemCode = line.ItemCode;
                            if (!String.IsNullOrWhiteSpace(line.ItemName)) docLine.ItemDescription = line.ItemName;
                            docLine.WarehouseCode = line.Whse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = line.WMSTransId;
                            docLine.Quantity = line.Quantity;

                            if (line.IsItemManagedBySnB(company) == "S")
                            {
                                var docLineSerial = docLine.SerialNumbers;
                                docLineSerial.BaseLineNumber = intLineCount;
                                //docLineSerial.InternalSerialNumber = line.SNBCode;

                                docLineSerial.SystemSerialNumber = SBOSupport.GetSerialSysNumber(company, line.ItemCode, line.SNBCode);
                                docLineSerial.Quantity = line.Quantity;
                                if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.ItemCode}'. Serial Number '{line.SNBCode}'. Total quantity must be 1.");
                            }
                            else if (line.IsItemManagedBySnB(company) == "B")
                            {
                                //managed by batch
                                var docLineBatch = docLine.BatchNumbers;
                                docLineBatch.BaseLineNumber = intLineCount;
                                docLineBatch.BatchNumber = line.SNBCode;
                                docLineBatch.Quantity = line.Quantity;
                            }

                            //Bin
                            if (line.IsBinWarehouse(company))
                            {
                                var docLineBin = docLine.BinAllocations;
                                docLineBin.BaseLineNumber = intLineCount;
                                docLineBin.BinAbsEntry = line.GetBinEntry(company);
                                docLineBin.Quantity = line.Quantity;
                                if (line.IsItemManagedBySnB(company) != "")
                                    docLineBin.SerialAndBatchNumbersBaseLine = 0;
                            }

                            intLineCount += 1;
                        }

                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_NEG, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_NEG : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), ige.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(ige), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + ige.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Stock Adjustment (-) failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_NEG : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), ige.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(ige), response.Remark, response.Status);

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

        public PostObjectResult POSTDOReturn(DOReturn rdn, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "RDN";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 JOIN {objectTableName}1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.Canceled='N' and T1.\"BaseType\" = -1 AND isnull(T0.\"U_AXC_EXTID\",'') = '{ rdn.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{rdn.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        rdn.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oReturns) as SAPbobsCOM.Documents;

                        HEADEROBJ.CardCode = rdn.CardCode;
                        HEADEROBJ.DocDueDate = rdn.GetDeliveryDate();
                        HEADEROBJ.DocDate = rdn.GetPostDate();
                        HEADEROBJ.Comments = rdn.Remark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = rdn.WMSTransId;

                        //Create one for each line
                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in rdn.Lines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            docLine.ItemCode = line.ItemCode;
                            if (!String.IsNullOrWhiteSpace(line.ItemName)) docLine.ItemDescription = line.ItemName;
                            docLine.WarehouseCode = line.Whse;
                            docLine.FreeText = line.ReturnReason ?? "";
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = line.WMSTransId;
                            docLine.Quantity = line.Quantity;
                            docLine.UseBaseUnits = line.IsInventoryUOM(company, out _)? BoYesNoEnum.tYES: BoYesNoEnum.tNO;

                            if (line.IsItemManagedBySnB(company) == "S")
                            {
                                var docLineSerial = docLine.SerialNumbers;
                                docLineSerial.BaseLineNumber = intLineCount;
                                docLineSerial.InternalSerialNumber = line.SNBCode;
                                docLineSerial.Quantity = line.Quantity;
                                if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.ItemCode}'. Serial Number '{line.SNBCode}'. Total quantity must be 1.");
                            }
                            else if (line.IsItemManagedBySnB(company) == "B")
                            {
                                //managed by batch
                                var docLineBatch = docLine.BatchNumbers;
                                docLineBatch.BaseLineNumber = intLineCount;
                                docLineBatch.BatchNumber = line.SNBCode;
                                docLineBatch.Quantity = (line.Quantity * line.GetUomQtyConversion(company)).RoundSAPAmount(company, null, RoundingContextEnum.rcQuantity);
                            }

                            //Bin
                            if (line.IsBinWarehouse(company))
                            {
                                var docLineBin = docLine.BinAllocations;
                                docLineBin.BaseLineNumber = intLineCount;
                                docLineBin.BinAbsEntry = line.GetBinEntry(company);
                                docLineBin.Quantity = (line.Quantity * line.GetUomQtyConversion(company)).RoundSAPAmount(company, null, RoundingContextEnum.rcQuantity);
                                if (line.IsItemManagedBySnB(company) != "")
                                    docLineBin.SerialAndBatchNumbersBaseLine = 0;
                            }

                            intLineCount += 1;
                        }

                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_AR_RETURN, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_AR_RETURN : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), rdn.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(rdn), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + rdn.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting DO Return (-) failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_AR_RETURN : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), rdn.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(rdn), response.Remark, response.Status);
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

        public PostObjectResult POSTDeliveryOrder(DeliveryOrder dln, string SQLConnection)
        {
            SAPbobsCOM.Documents HEADEROBJ = null;
            string objectTableName = "DLN";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 where T0.Canceled='N' AND isnull(T0.\"U_AXC_EXTID\",'') = '{ dln.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{dln.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        dln.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oDeliveryNotes) as SAPbobsCOM.Documents;
                        HEADEROBJ.CardCode = dln.CardCode;
                        HEADEROBJ.DocDate = dln.GetPostDate();
                        HEADEROBJ.DocDueDate = dln.GetDeliveryDate();
                        HEADEROBJ.Comments = dln.Remark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = dln.WMSTransId;

                        //Create one for each line
                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in dln.Lines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            if (!String.IsNullOrWhiteSpace(line.BaseWMSTransId))
                            {
                                int baseEntry = line.GetBaseEntry(company, dln.BaseWMSTransId);
                                int baseLine = line.GetBaseLine(company, dln.BaseWMSTransId);
                                if (baseEntry == 0) throw new Exception($"Line [{line.WMSTransId}]. Base Entry ({line.BaseWMSTransId}) not found.");
                                docLine.BaseType = (int)SAPbobsCOM.BoObjectTypes.oReturns;
                                docLine.BaseEntry = baseEntry;
                                docLine.BaseLine = baseLine;
                            }
                            docLine.ItemCode = line.ItemCode;
                            if (!String.IsNullOrWhiteSpace(line.ItemName)) docLine.ItemDescription = line.ItemName;
                            docLine.WarehouseCode = line.Whse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = line.WMSTransId;
                            docLine.Quantity = line.Quantity;

                            if (line.IsItemManagedBySnB(company) == "S")
                            {
                                var docLineSerial = docLine.SerialNumbers;
                                docLineSerial.BaseLineNumber = intLineCount;
                                docLineSerial.InternalSerialNumber = line.SNBCode;
                                docLineSerial.Quantity = line.Quantity;
                                if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.ItemCode}'. Serial Number '{line.SNBCode}'. Total quantity must be 1.");

                            }
                            else if (line.IsItemManagedBySnB(company) == "B")
                            {
                                //managed by batch
                                var docLineBatch = docLine.BatchNumbers;
                                docLineBatch.BaseLineNumber = intLineCount;
                                docLineBatch.BatchNumber = line.SNBCode;
                                docLineBatch.Quantity = line.Quantity;
                            }

                            //Bin
                            if (line.IsBinWarehouse(company))
                            {
                                var docLineBin = docLine.BinAllocations;
                                docLineBin.BaseLineNumber = intLineCount;
                                docLineBin.BinAbsEntry = line.GetBinEntry(company);
                                docLineBin.Quantity = line.Quantity;
                                if (line.IsItemManagedBySnB(company) != "")
                                    docLineBin.SerialAndBatchNumbersBaseLine = 0;
                            }

                            intLineCount += 1;
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_DELIVERY_ORDER, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_DELIVERY_ORDER : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), dln.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(dln), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + dln.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Delivery Order failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_DELIVERY_ORDER : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), dln.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(dln), response.Remark, response.Status);

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

        public PostObjectResult POSTInventoryTrsReq(InventoryTrReq wtq, string SQLConnection)
        {
            SAPbobsCOM.StockTransfer HEADEROBJ = null;
            string objectTableName = "WTQ";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0  where T0.Canceled='N' AND isnull(T0.\"U_AXC_EXTID\",'') = '{ wtq.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{wtq.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        wtq.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oInventoryTransferRequest) as SAPbobsCOM.StockTransfer;

                        HEADEROBJ.DocDate = wtq.GetPostDate();
                        HEADEROBJ.Comments = wtq.Remark;
                        if (!String.IsNullOrWhiteSpace(wtq.JournalRemark)) HEADEROBJ.JournalMemo = wtq.JournalRemark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = wtq.WMSTransId;
                        HEADEROBJ.FromWarehouse = wtq.Lines.Select(x => x.FromWhse).FirstOrDefault();
                        HEADEROBJ.ToWarehouse = wtq.Lines.Select(x => x.ToWhse).FirstOrDefault();

                        //Create one for each line
                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in wtq.Lines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            docLine.ItemCode = line.ItemCode;
                            if(!String.IsNullOrWhiteSpace(line.ItemName)) docLine.ItemDescription = line.ItemName;
                            docLine.WarehouseCode = line.ToWhse;
                            docLine.FromWarehouseCode = line.FromWhse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = line.WMSTransId;
                            docLine.Quantity = line.Quantity;

                            if (line.IsItemManagedBySnB(company) == "S")
                            {
                                var docLineSerial = docLine.SerialNumbers;
                                docLineSerial.BaseLineNumber = intLineCount;
                                docLineSerial.InternalSerialNumber = line.SNBCode;
                                docLineSerial.Quantity = line.Quantity;
                                if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.ItemCode}'. Serial Number '{line.SNBCode}'. Total quantity must be 1.");

                                //No Bin in Transfer Reqquest
                            }
                            else if (line.IsItemManagedBySnB(company) == "B")
                            {
                                //managed by batch
                                var docLineBatch = docLine.BatchNumbers;
                                docLineBatch.BaseLineNumber = intLineCount;
                                docLineBatch.BatchNumber = line.SNBCode;
                                docLineBatch.Quantity = line.Quantity;
                            }

                            intLineCount += 1;
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_TR_REQUEST, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_TR_REQUEST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), wtq.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(wtq), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + wtq.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Stock Transfer Request failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_TR_REQUEST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), wtq.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(wtq), response.Remark, response.Status);

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

        public PostObjectResult POSTInventoryTransfer(InventoryTransfer wtr, string SQLConnection)
        {
            SAPbobsCOM.StockTransfer HEADEROBJ = null;
            string objectTableName = "WTR";
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from O{objectTableName} T0 where T0.Canceled='N' AND isnull(T0.\"U_AXC_EXTID\",'') = '{ wtr.WMSTransId}'", company));

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{wtr.WMSTransId}' already Exists in SAP Business One! [O{objectTableName}]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        //Validate the lines first.
                        wtr.ValidateLine(company);
                        HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oStockTransfer) as SAPbobsCOM.StockTransfer;

                        HEADEROBJ.DocDate = wtr.GetPostDate();
                        HEADEROBJ.Comments = wtr.Remark;
                        if (!String.IsNullOrWhiteSpace(wtr.JournalRemark)) HEADEROBJ.JournalMemo = wtr.JournalRemark;
                        HEADEROBJ.UserFields.Fields.Item("U_AXC_EXTID").Value = wtr.WMSTransId;
                        HEADEROBJ.FromWarehouse = wtr.Lines.Select(x => x.FromWhse).FirstOrDefault();
                        HEADEROBJ.ToWarehouse = wtr.Lines.Select(x => x.ToWhse).FirstOrDefault();

                        //Create one for each line
                        int intLineCount = 0;
                        var docLine = HEADEROBJ.Lines;
                        foreach (var line in wtr.Lines)
                        {
                            if (intLineCount > 0)
                                docLine.Add();

                            if(!String.IsNullOrWhiteSpace(line.RqWMSTransId))
                            {
                                int baseEntry = line.GetBaseEntry(company, wtr.RqWMSTransId);
                                int baseLine = line.GetBaseLine(company, wtr.RqWMSTransId);
                                if (baseEntry == 0) throw new Exception($"Line [{line.WMSTransId}]. Base Entry ({line.RqWMSTransId}) not found.");
                                docLine.BaseType = InvBaseDocTypeEnum.InventoryTransferRequest;
                                docLine.BaseEntry = baseEntry;
                                docLine.BaseLine = baseLine;
                            }
                            docLine.ItemCode = line.ItemCode;
                            if (!String.IsNullOrWhiteSpace(line.ItemName)) docLine.ItemDescription = line.ItemName;
                            docLine.WarehouseCode = line.ToWhse;
                            docLine.FromWarehouseCode = line.FromWhse;
                            docLine.UserFields.Fields.Item("U_AXC_EXTID").Value = line.WMSTransId;
                            docLine.Quantity = line.Quantity;

                            if (line.IsItemManagedBySnB(company) == "S")
                            {
                                var docLineSerial = docLine.SerialNumbers;
                                docLineSerial.BaseLineNumber = intLineCount;
                                docLineSerial.InternalSerialNumber = line.SNBCode;
                                docLineSerial.Quantity = line.Quantity;
                                if (docLineSerial.Quantity > 1) throw new Exception($"Item '{line.ItemCode}'. Serial Number '{line.SNBCode}'. Total quantity must be 1.");

                            }
                            else if (line.IsItemManagedBySnB(company) == "B")
                            {
                                //managed by batch
                                var docLineBatch = docLine.BatchNumbers;
                                docLineBatch.BaseLineNumber = intLineCount;
                                docLineBatch.BatchNumber = line.SNBCode;
                                docLineBatch.Quantity = line.Quantity;
                            }

                            //Bin
                            var docLineBin = docLine.BinAllocations;
                            int binLineCount = 0;
                            if (line.IsFromBinWarehouse(company))
                            {
                                binLineCount++;
                                docLineBin.BaseLineNumber = intLineCount;
                                docLineBin.BinActionType = BinActionTypeEnum.batFromWarehouse;
                                docLineBin.BinAbsEntry = line.GetFromBinEntry(company);
                                docLineBin.Quantity = line.Quantity;
                                if (line.IsItemManagedBySnB(company) != "")
                                    docLineBin.SerialAndBatchNumbersBaseLine = 0;
                            }
                            if (line.IsToBinWarehouse(company))
                            {
                                if (binLineCount > 0)
                                    docLineBin.Add();
                                docLineBin.BaseLineNumber = intLineCount;
                                docLineBin.BinActionType = BinActionTypeEnum.batToWarehouse;
                                docLineBin.BinAbsEntry = line.GetToBinEntry(company);
                                docLineBin.Quantity = line.Quantity;
                                if (line.IsItemManagedBySnB(company) != "")
                                    docLineBin.SerialAndBatchNumbersBaseLine = 0;
                            }


                            intLineCount += 1;
                        }
                        string pickListString2 = HEADEROBJ.GetAsXML();

                        lerrCode = HEADEROBJ.Add();
                        if (lerrCode == 0)
                        {
                            try
                            {
                                ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                string objType = company.GetNewObjectType();
                                string tableName = SBOSupport.GetTableName((SAPbobsCOM.BoObjectTypes)int.Parse(objType));
                                lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE($"select \"DocNum\" from O{tableName} where \"DocEntry\" = '{ValDocEntry}'", company));
                                response = new PostObjectResult();
                                response.Status = true;
                                response.DocEntry = int.Parse(ValDocEntry);
                                response.DocNumber = response.DocEntry;

                                //Send the Alert
                                SBOSupport.SendAlert(company, true, PostObjectPayload.SYNCH_I_OBJECT_WHS_TRANSFER, objType, ValDocEntry, ValDocEntry);
                            }
                            catch
                            {
#if DEBUG
                                throw;
#endif
                            }

                            AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_WHS_TRANSFER : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), wtr.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(wtr), response.Remark, response.Status);

                        }
                        else
                        {
                            lastErrorMessage = "Transaction Id: " + wtr.WMSTransId + " - " + company.GetLastErrorDescription();
                            throw new Exception(lastErrorMessage);
                            ;
                        }
                    }

                    try
                    {
                        if (HEADEROBJ != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                        if (company != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch
                    { }
                }
                else
                {
                    throw new Exception($"Connection to Company fails! {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {
                response = new PostObjectResult($"Posting Stock Transfer failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_WHS_TRANSFER : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), wtr.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(wtr), response.Remark, response.Status);

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
    }
}
