using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SBOCLASS.Models;
using SAPbobsCOM;
namespace SBOCLASS.Class
{
    public class API_InvoiceClass
    {
        SAPbobsCOM.Company company { get; set; }
        SAPbobsCOM.Documents HEADEROBJ { get; set; }
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

        public List<Models.GetPaymentStatus> GetPaid(string SQLConnection, string TransID = "")
        {
            try
            {
                //Declarations
                List<Models.GetPaymentStatus> returnData = new List<Models.GetPaymentStatus>();
                System.Data.DataTable dt;

                SQLClass sql = new SQLClass();
                sql.ConnectionString = SQLConnection;

                sql.CommandType = System.Data.CommandType.Text;
                if (TransID == "")
                {
                    dt = sql.GetDataTable("select * from \"INV_GETPAYMENTSTATUS\"");
                }
                else
                {
                    dt = sql.GetDataTable("select * from \"INV_GETPAYMENTSTATUS\" where \"U_TransId\" = '" + TransID + "'");
                }

                //Get Header Data
                returnData = (from System.Data.DataRow row in dt.Rows

                              select new Models.GetPaymentStatus
                              {
                                  U_TransId = row[0].ToString(),
                                  Status = row[1].ToString(),
                                  DocTotal = Convert.ToDouble(row[2].ToString()),
                                  AppliedAmount = Convert.ToDouble(row[3].ToString()),
                                  BalanceDue = Convert.ToDouble(row[4].ToString()),
                                  PaidDate = row[5].ToString(),
                                  Currency = row[6].ToString()
                              }).ToList();


                return returnData;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public ResponseResult POSTInvoice(API_InvoiceClassHeader InvHead, string SQLConnection)
        {
            string ProductID = string.Empty;
            string ValDocEntry = string.Empty;
            string VATCODE = string.Empty;
            ResponseResult response = new ResponseResult();
            int intCheckID_If_Exist;
            try
            {
                try
                {
                    if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                    {
                        int intLineCount = 0;
                        intCheckID_If_Exist = Convert.ToInt32(SBOSupport.GETSINGLEVALUE("select count(\"DocEntry\") from \"OINV\" where isnull(\"U_TransId\",'') = '" + InvHead.U_TransId + "'", company));

                        if (intCheckID_If_Exist > 0)
                        {
                            lastErrorMessage = "Portal Transaction ID: " + Convert.ToString(InvHead.U_TransId) + " Already Exists in SAP Business One!";
                            response = new ResponseResult();
                            response.RecordStatus = "true";
                            response.ErrorDescription = lastErrorMessage;
                        }
                        else
                        {
                            HEADEROBJ = company.GetBusinessObject(BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;

                            HEADEROBJ.CardCode = InvHead.CardCode;
                            HEADEROBJ.DocDate = InvHead.DocDate;
                            HEADEROBJ.NumAtCard = Convert.ToString(InvHead.NumAtCard);
                            HEADEROBJ.UserFields.Fields.Item("U_TransId").Value = InvHead.U_TransId;

                            foreach (var dat in InvHead.Details.ToList())
                            {
                                intLineCount += 1;
                                HEADEROBJ.Lines.ItemCode = dat.ItemCode;
                                HEADEROBJ.Lines.Price = dat.Price;
                                HEADEROBJ.Lines.Quantity = dat.Quantity;

                                HEADEROBJ.Lines.Add();
                            }

                            lerrCode = HEADEROBJ.Add();

                            if (lerrCode == 0)
                            {
                                try
                                {
                                    ValDocEntry = Convert.ToString(company.GetNewObjectKey());
                                    lastErrorMessage = Convert.ToString(SBOSupport.GETSINGLEVALUE("select \"DocNum\" from \"OINV\" where \"DocEntry\" = '" + ValDocEntry + "'", company));
                                    response = new ResponseResult();
                                    response.RecordStatus = "true";
                                    response.ErrorDescription = lastErrorMessage;
                                }
                                catch
                                { }
                            }
                            else
                            {
                                lastErrorMessage = "Transaction Id: " + InvHead.U_TransId + " - " + company.GetLastErrorDescription();
                                response = new ResponseResult();
                                response.RecordStatus = "false";
                                response.ErrorDescription = lastErrorMessage;
                            }
                        }

                        try
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(HEADEROBJ);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                        }
                        catch
                        { }
                    }
                    else
                    {
                        response = new ResponseResult();
                        response.RecordStatus = "false";
                        response.ErrorDescription = "Connection to Company fails! " + lastErrorMessage;

                        try
                        { System.Runtime.InteropServices.Marshal.ReleaseComObject(company); }
                        catch
                        { }
                    }
                }
                catch (Exception ex)
                {
                    response = new ResponseResult();
                    response.RecordStatus = "false";
                    response.ErrorDescription = "Exception Error Method Posting: " + ex.Message;
                    try
                    { System.Runtime.InteropServices.Marshal.ReleaseComObject(company); }
                    catch
                    { }
                }
            }
            catch (Exception ex)
            {
                response = new ResponseResult();
                response.RecordStatus = "false";
                response.ErrorDescription = "Exception Error Method Posting: " + ex.Message;
                try
                { System.Runtime.InteropServices.Marshal.ReleaseComObject(company); }
                catch
                { }
            }

            return response;
        }
    }
}
