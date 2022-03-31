using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SBOCLASS.Models.EOA;
using SAPbobsCOM;

namespace SBOCLASS.Class
{
    public class API_StockPostingClass
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

        public PostObjectResult POSTStockPosting(StockPosting stockpost, string SQLConnection)
        {
            string ValDocEntry = string.Empty;
            PostObjectResult response = new PostObjectResult();
            int intCheckID_If_Exist;
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from \"OIQR\" where isnull(\"U_AXC_EXTID\",'') = '{ stockpost.WMSTransId}'", company));
                    stockpost.ValidateLine(company);

                    if (intCheckID_If_Exist > 0)
                    {
                        lastErrorMessage = $"WMS Transaction ID: '{stockpost.WMSTransId}' Already Exists in SAP Business One! [OIQR]";
                        response = new PostObjectResult(lastErrorMessage, true);
                    }
                    else
                    {
                        var cntLines = stockpost.Lines
                           .GroupBy(x => new {
                               x.BaseKey,
                               x.BaseLine,
                               BaseType = x.GetSAPBaseType(),
                               x.Whse,
                               IsBinWarehouse = x.IsBinWarehouse(company),
                               x.ItemCode,
                               IsItemManagedBySnB = x.IsItemManagedBySnB(company),
                               x.CountQty,
                               binEntry = x.GetBinEntry(company),
                               Uom = x.UOM,
                               SNB = x.SNBCode,
                               ExRef = x.WMSTransId,
                               Line = x.LineNo
                           }).Select(x => x);

                        CompanyService oCS = (SAPbobsCOM.CompanyService)company.GetCompanyService();
                        try
                        {
                            //CompanyService oCS = (SAPbobsCOM.CompanyService)company.GetCompanyService();
                            SAPbobsCOM.InventoryCountingsService oInventoryCountingsService = (InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                            InventoryCounting oInventoryCounting = (InventoryCounting)oInventoryCountingsService.GetDataInterface(InventoryCountingsServiceDataInterfaces.icsInventoryCounting);
                            SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oInventoryCountingsService.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                            SAPbobsCOM.InventoryCountingLine oICline;

                            int docEntry = 0;
                            foreach (var line in cntLines)
                            {
                                if (docEntry == 0)
                                {
                                    docEntry = line.Key.BaseKey;
                                    oICP.DocumentEntry = docEntry;
                                    oInventoryCounting = oInventoryCountingsService.Get(oICP);

                                    oICline = oInventoryCounting.InventoryCountingLines.Item(line.Key.BaseLine -1);
                                    oICline.Counted = BoYesNoEnum.tYES;
                                    oICline.CountedQuantity = line.Key.CountQty;
                                }
                                else
                                {
                                    oICline = oInventoryCounting.InventoryCountingLines.Item(line.Key.BaseLine -1);
                                    oICline.Counted = BoYesNoEnum.tYES;
                                    oICline.CountedQuantity = line.Key.CountQty;
                                }
                            }
                            oInventoryCountingsService.Update(oInventoryCounting);
                            

                            InventoryPostingsService oInventoryPostingsService = (InventoryPostingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryPostingsService);
                            InventoryPosting oInventoryPosting = (InventoryPosting)oInventoryPostingsService.GetDataInterface(InventoryPostingsServiceDataInterfaces.ipsInventoryPosting);
                            oInventoryPosting.CountDate = stockpost.GetCountDate();
                            oInventoryPosting.Remarks = stockpost.Remark;
                            oInventoryPosting.UserFields.Item("U_AXC_EXTID").Value = stockpost.WMSTransId;
                            //oInventoryPosting.PriceList = 1;
                            InventoryPostingLines oInventoryPostingLines = oInventoryPosting.InventoryPostingLines;

                            int intItmMan_If_Exist = 0;
                            foreach (var line in cntLines)
                            {                                
                                InventoryPostingLine oInventoryPostingLine = oInventoryPostingLines.Add();
                                oInventoryPostingLine.ItemCode = line.Key.ItemCode;
                                oInventoryPostingLine.CountedQuantity = line.Key.CountQty;
                                oInventoryPostingLine.WarehouseCode = line.Key.Whse;
                                oInventoryPostingLine.BinEntry = line.Key.binEntry;
                                oInventoryPostingLine.BaseEntry = line.Key.BaseKey;
                                oInventoryPostingLine.BaseType = line.Key.BaseType;
                                oInventoryPostingLine.BaseLine = line.Key.BaseLine;
                                oInventoryPostingLine.CountDate = stockpost.GetCountDate();
                                oInventoryPostingLine.UserFields.Item("U_AXC_EXTID").Value = line.Key.ExRef;

                                intItmMan_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from \"OITM\" where \"ManBtchNum\" = 'Y' and \"ItemCode\" = '{line.Key.ItemCode}'", company));

                                if (!String.IsNullOrEmpty(line.Key.SNB.ToString().Trim()) && intItmMan_If_Exist>0)
                                {
                                    InventoryPostingBatchNumber oInventoryPostingBatchNumber = oInventoryPostingLine.InventoryPostingBatchNumbers.Add();
                                    oInventoryPostingBatchNumber.BatchNumber = line.Key.SNB;
                                    oInventoryPostingBatchNumber.Quantity = line.Key.CountQty;
                                }
                                intItmMan_If_Exist = 0;
                                intItmMan_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE($"select count('') from \"OITM\" where \"ManSerNum\" = 'Y' and \"ItemCode\" = '{line.Key.ItemCode}'", company));

                                if (!String.IsNullOrEmpty(line.Key.SNB.ToString().Trim()) && intItmMan_If_Exist > 0)
                                {
                                    InventoryPostingSerialNumber oInventoryPostingSerial = oInventoryPostingLine.InventoryPostingSerialNumbers.Add();
                                    oInventoryPostingSerial.ManufacturerSerialNumber = line.Key.SNB;
                                    oInventoryPostingSerial.Quantity = line.Key.CountQty;
                                }
                            }
                            InventoryPostingParams oInventoryPostingParams = oInventoryPostingsService.Add(oInventoryPosting);
                            response.ObjType = SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_STOCK_POST;
                            response.DocEntry = oInventoryPostingParams.DocumentEntry;
                            response.DocNumber = oInventoryPostingParams.DocumentNumber;
                            response.Remark = "Posted Successfully";
                        }
                        catch (Exception ex)
                        {
                            response = new PostObjectResult($"Posting Inventory Posting failed: {ex.Message}");
                            if (company != null && company.Connected && company.InTransaction)
                                company.EndTransaction(BoWfTransOpt.wf_RollBack);
                                AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_POST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), stockpost.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(stockpost), response.Remark, response.Status);
                        }
                        
                    }

                    AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_POST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), stockpost.WMSTransId, SBOSupport.Operation.POST,
                                Newtonsoft.Json.JsonConvert.SerializeObject(stockpost), response.Remark, response.Status);

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
                response = new PostObjectResult($"Posting Inventory Posting failed: {ex.Message}");
                if (company != null && company.Connected && company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                    AXC_OFTLG.GenerateLogRecord(company, (response.ObjType ?? "") == "" ? PostObjectPayload.SYNCH_I_OBJECT_STOCK_POST : response.ObjType, response.DocEntry == 0 ? "" : response.DocEntry.ToString(), response.DocNumber == 0 ? "" : response.DocNumber.ToString(), stockpost.WMSTransId, SBOSupport.Operation.POST,
                    Newtonsoft.Json.JsonConvert.SerializeObject(stockpost), response.Remark, response.Status);
            }
            finally
            {
                if (company == null && company.InTransaction) company.EndTransaction(BoWfTransOpt.wf_RollBack);
                try
                { System.Runtime.InteropServices.Marshal.ReleaseComObject(company); }
                catch
                { }
            }
            response.Status = response.DocEntry == 0 ? false : true;
            return response;
        }
    }
}
                                //if (docEntry != 0)
                                //{
                                //    if(docEntry==line.Key.BaseKey)
                                //    {
                                //        oICline = oInventoryCounting.InventoryCountingLines.Item(line.Key.BaseLine);
                                //        oICline.Counted = BoYesNoEnum.tYES;
                                //        oICline.CountedQuantity = line.Key.CountQty;
                                //    }
                                //    else
                                //    {

                                //        oInventoryCountingsService.Update(oInventoryCounting);

                                //        docEntry = line.Key.BaseKey;
                                //        oICP.DocumentEntry = docEntry;
                                //        oInventoryCounting = oInventoryCountingsService.Get(oICP);

                                //        oICline = oInventoryCounting.InventoryCountingLines.Item(line.Key.BaseLine);
                                //        oICline.Counted = BoYesNoEnum.tYES;
                                //        oICline.CountedQuantity = line.Key.CountQty;
                                //    }
                                //}
                                //else
                                //{
                                //    docEntry = line.Key.BaseKey;
                                //    oICP.DocumentEntry = docEntry;
                                //    oInventoryCounting = oInventoryCountingsService.Get(oICP);

                                //    oICline = oInventoryCounting.InventoryCountingLines.Item(line.Key.BaseLine);
                                //    oICline.Counted = BoYesNoEnum.tYES;
                                //    oICline.CountedQuantity = line.Key.CountQty;
                                //}
