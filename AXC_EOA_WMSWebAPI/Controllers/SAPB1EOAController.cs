using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using SBOCLASS.Class;
using SBOCLASS.Interface;
using SBOCLASS.Models;
using SBOCLASS.Models.EOA;

namespace AXC_EOA_WMSWebAPI.Controllers
{
    public class SAPB1EOAController : ApiController
    {
        [Route("")]
        [AXC_EOA_WMSWebAPI.Authenticate.IdentityBasicAuthentication]
        [System.Web.Mvc.HttpPost]
        public HttpResponseMessage SAPObject(SBOCLASS.Models.EOA.PostObjectPayload payload)
        {
            HttpResponseMessage response = null;
            SBOCLASS.Models.EOA.PostObjectResult returnResult = new PostObjectResult();
            try
            {
                //Declarations

                SetValuestoConstants();
                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["SAPConnString"].ToString();
                API_DocumentClass docClass;                
                if (payload?.Data == null)
                    throw new MissingFieldException("Missing payload data");


                switch (payload.ObjType)
                {
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_PICK_LIST:
                        //Create PickList - Basedon SO or Reserve Invoice
                        PickList pickList = Newtonsoft.Json.JsonConvert.DeserializeObject<PickList>(payload.Data.ToString());
                        if (pickList == null)
                            throw new MissingFieldException("Invalid picklist payload data");
                        pickList.Validate();
                        API_PickListClass pickListClass = new API_PickListClass();
                        returnResult = pickListClass.POSTPickList(pickList, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_GRPO:
                        switch (payload.BaseObj ?? "")
                        {
                            case PostObjectPayload.SYNCH_O_OBJECT_PURCHASE_ORDER:
                            case "":
                                GRPO grpo = Newtonsoft.Json.JsonConvert.DeserializeObject<GRPO>(payload.Data.ToString());
                                if (grpo == null)
                                    throw new MissingFieldException("Invalid GRPO payload data");
                                grpo.Validate();
                                docClass = new API_DocumentClass();
                                returnResult = docClass.POSTGRPO(grpo, connStr);
                                break;
                            default:
                                returnResult = new PostObjectResult($"Invalid base type for object ({payload.ObjType})GRPO. Valid values are ({PostObjectPayload.SYNCH_O_OBJECT_PURCHASE_ORDER})PO.");
                                break;
                        }
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_ISSUE_PROD:
                        //Create Issue for production
                        IssueForProduction igeP = Newtonsoft.Json.JsonConvert.DeserializeObject<IssueForProduction>(payload.Data.ToString());
                        if (igeP == null)
                            throw new MissingFieldException("Invalid IssueForProduction payload data.");
                        igeP.Validate();
                        docClass = new API_DocumentClass();
                        returnResult = docClass.POSTIssueProd(igeP, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_RECPT_PROD:
                        //Create Issue for production
                        ReceiptFromProduction ignP = Newtonsoft.Json.JsonConvert.DeserializeObject<ReceiptFromProduction>(payload.Data.ToString());
                        if (ignP == null)
                            throw new MissingFieldException("Invalid ReceiptFromProduction payload data.");
                        ignP.Validate();
                        docClass = new API_DocumentClass();
                        returnResult = docClass.POSTReceiptProd(ignP, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_NEG:
                        //Create Goods Issue
                        StockAdjustmentNeg ige = Newtonsoft.Json.JsonConvert.DeserializeObject<StockAdjustmentNeg>(payload.Data.ToString());
                        if (ige == null)
                            throw new MissingFieldException("Invalid StockAdjusment payload data.");
                        docClass = new API_DocumentClass();
                        ige.Validate();
                        returnResult = docClass.POSTStockAdjNeg(ige, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_STOCK_ADJ_POS:
                        //Create Goods Receipt
                        StockAdjustmentPos ign = Newtonsoft.Json.JsonConvert.DeserializeObject<StockAdjustmentPos>(payload.Data.ToString());
                        if (ign == null)
                            throw new MissingFieldException("Invalid StockAdjusment payload data.");
                        ign.Validate();
                        docClass = new API_DocumentClass();
                        returnResult = docClass.POSTStockAdjPos(ign, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_TR_REQUEST:
                        //Create Goods Issue
                        InventoryTrReq wtq = Newtonsoft.Json.JsonConvert.DeserializeObject<InventoryTrReq>(payload.Data.ToString());
                        if (wtq == null)
                            throw new MissingFieldException("Invalid InventoryTransferRequest payload data.");
                        docClass = new API_DocumentClass();
                        wtq.Validate();
                        returnResult = docClass.POSTInventoryTrsReq(wtq, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_WHS_TRANSFER:
                        //Create Goods Issue
                        InventoryTransfer wtr = Newtonsoft.Json.JsonConvert.DeserializeObject<InventoryTransfer>(payload.Data.ToString());
                        if (wtr == null)
                            throw new MissingFieldException("Invalid InventoryTransfer payload data.");
                        wtr.Validate();
                        docClass = new API_DocumentClass();
                        returnResult = docClass.POSTInventoryTransfer(wtr, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_AR_RETURN:
                        //Create DO Return
                        DOReturn rdn = Newtonsoft.Json.JsonConvert.DeserializeObject<DOReturn>(payload.Data.ToString());
                        if (rdn == null)
                            throw new MissingFieldException("Invalid DO Return payload data.");
                        rdn.Validate();
                        docClass = new API_DocumentClass();
                        returnResult = docClass.POSTDOReturn(rdn, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_DELIVERY_ORDER:
                        //Create DO 
                        DeliveryOrder dln = Newtonsoft.Json.JsonConvert.DeserializeObject<DeliveryOrder>(payload.Data.ToString());
                        if (dln == null)
                            throw new MissingFieldException("Invalid DO payload data.");
                        dln.Validate();
                        docClass = new API_DocumentClass();
                        returnResult = docClass.POSTDeliveryOrder(dln, connStr);
                        break;
                    case SBOCLASS.Models.EOA.PostObjectPayload.SYNCH_I_OBJECT_STOCK_POST:
                        //Create Inventory Posting
                        StockPosting stk = Newtonsoft.Json.JsonConvert.DeserializeObject<StockPosting>(payload.Data.ToString());
                        if (stk == null)
                            throw new MissingFieldException("Invalid Stock Posting payload data.");
                        stk.Validate();
                        API_StockPostingClass StkClass = new API_StockPostingClass();
                        returnResult = StkClass.POSTStockPosting(stk, connStr);
                        break;
                }

                //Convert result response
                string jsonResult = Newtonsoft.Json.JsonConvert.SerializeObject(returnResult);
                if (!string.IsNullOrEmpty(jsonResult))
                {
                    response = Request.CreateResponse<PostObjectResult>(HttpStatusCode.Created, returnResult);
                }
            }
            catch(MissingFieldException mex)
            {
                returnResult = new PostObjectResult(mex.Message);
                response = Request.CreateResponse<PostObjectResult>(HttpStatusCode.BadRequest, returnResult);
            }
            catch (Exception ex)
            {
                returnResult = new PostObjectResult(ex.Message);
                response = Request.CreateResponse<PostObjectResult>(HttpStatusCode.InternalServerError, returnResult);
                //throw ex;
            }
            return response;

        }


        [AXC_EOA_WMSWebAPI.Authenticate.IdentityBasicAuthentication]
        [System.Web.Mvc.HttpPost]
        public HttpResponseMessage POST_INVOICE(API_InvoiceClassHeader Invoice)
        {
            try
            {
                //Declarations
                API_InvoiceClassHeader head = Invoice;
                ResponseResult returnResult;
                HttpResponseMessage response = null;


                SetValuestoConstants();
                API_InvoiceClass inv = new API_InvoiceClass();

                string jsonResult = string.Empty;
                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["SAPConnString"].ToString();
                returnResult = inv.POSTInvoice(head, connStr);

                //Convert result response
                jsonResult = Newtonsoft.Json.JsonConvert.SerializeObject(returnResult);

                if (!string.IsNullOrEmpty(jsonResult))
                {
                    response = Request.CreateResponse<ResponseResult>(HttpStatusCode.Created, returnResult);
                }
                return response;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [AXC_EOA_WMSWebAPI.Authenticate.IdentityBasicAuthentication]
        [System.Web.Mvc.HttpGet]
        public List<GetPaymentStatus> Get_Paid(string TransId = "")
        {
            try
            {
                //Declarations
                List<GetPaymentStatus> lst;
                API_InvoiceClass master = new API_InvoiceClass();

                string strReturn = string.Empty;

                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["SAPConnString"].ToString();
                if (TransId == "")
                {
                    lst = master.GetPaid(connStr);
                }
                else
                {
                    lst = master.GetPaid(connStr,TransId);
                }
                strReturn = Newtonsoft.Json.JsonConvert.SerializeObject(lst);
                //return strReturn;
                return lst;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetValuestoConstants()
        {
            try
            {
                SBOCLASS.Class.ConstantClass.Database = AXC_EOA_WMSWebAPI.Class.ConstantClass.Database;
                SBOCLASS.Class.ConstantClass.SBOServer = AXC_EOA_WMSWebAPI.Class.ConstantClass.SBOServer;
                SBOCLASS.Class.ConstantClass.SQLVersion = AXC_EOA_WMSWebAPI.Class.ConstantClass.SQLVersion;
                SBOCLASS.Class.ConstantClass.SAPUser = AXC_EOA_WMSWebAPI.Class.ConstantClass.SAPUser;
                SBOCLASS.Class.ConstantClass.SAPPassword = AXC_EOA_WMSWebAPI.Class.ConstantClass.SAPPassword;
                SBOCLASS.Class.ConstantClass.SQLUserName = AXC_EOA_WMSWebAPI.Class.ConstantClass.SQLUserName;
                SBOCLASS.Class.ConstantClass.SQLPassword = AXC_EOA_WMSWebAPI.Class.ConstantClass.SQLPassword;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }   
    }
}
