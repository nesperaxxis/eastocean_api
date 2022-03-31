using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using SAPbobsCOM;

namespace SBOCLASS.Class
{
    public class SBOClass
    {

        SAPbobsCOM.Company sapCompany;
        string lastErrorMessage = string.Empty;
        string lastErrorMessage_Out = string.Empty;
        string CostCode = string.Empty;
        string GLAccount = string.Empty;
        string VATCode = string.Empty;

        public bool connectToLoginCompany(string SQLServerName, string CompanyDB, string DBUserName, string DBPassword, string SBOUserName, string SBOPassword)
        {
            bool functionReturnValue = false;

            int lErrCode = 0;

            try
            {
                //// Initialize the Company Object.
                //// Create a new company object
                sapCompany = new SAPbobsCOM.Company();

                //// Set the mandatory properties for the connection to the database.
                //// To use a remote Db Server enter his name instead of the string "(local)"
                //// This string is used to work on a DB installed on your local machine

                sapCompany.Server = SQLServerName;
                sapCompany.CompanyDB = CompanyDB;
                sapCompany.UserName = SBOUserName;
                sapCompany.Password = SBOPassword;
                sapCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;

                //// Use Windows authentication for database server.
                //// True for NT server authentication,
                //// False for database server authentication.
                sapCompany.UseTrusted = false;
                if (ConstantClass.SQLVersion == 2012)
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                }
                else if (ConstantClass.SQLVersion == 2008)
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                }

                sapCompany.DbUserName = DBUserName;
                sapCompany.DbPassword = DBPassword;

                //// connect
                lErrCode = sapCompany.Connect();

                //// Check for errors during connect
                //sapCompany.GetLastError(lErrCode,lastErrorMessage_Out);
                if (lErrCode != 0)
                {
                    lastErrorMessage = "SAP Connection Error : " + sapCompany.GetLastErrorDescription();
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
        
        //Code for Creation of UDF 
        #region "CreateSettings"
        public string createUDF(String tableName, String fieldName, String fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int fieldSize, SAPbobsCOM.BoFldSubTypes subfieldType = SAPbobsCOM.BoFldSubTypes.st_None, String fieldValues = "", String defaultValue = "", String linkTable = null, string DBCompany = "")
        {
            //Declarations for SQLQuery

            try
            {
                string sqlScript = "select Top 1 fieldID from [" + sapCompany.CompanyDB + "].dbo.cufd where TableID = '" + tableName + "' and AliasID = '" + fieldName + "'";
                SAPbobsCOM.Recordset oRecset = (SAPbobsCOM.Recordset)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecset.DoQuery(sqlScript);

                //Execute Selected Query
                if (oRecset.RecordCount != 0)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset);
                    }
                    catch
                    {
                    }
                   return "UDF Already Exist!";
                }

                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset);
                }
                catch
                {
                }

                GC.Collect();
                SAPbobsCOM.UserFieldsMD oUDF = default(SAPbobsCOM.UserFieldsMD);
                oUDF = null;

                oUDF = (SAPbobsCOM.UserFieldsMD)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                //Filling userdefinefields data.
                oUDF.Name = fieldName;
                oUDF.Type = fieldType;
                oUDF.Size = fieldSize;
                oUDF.Description = fieldDescription;
                oUDF.TableName = tableName;
                oUDF.EditSize = fieldSize;
                oUDF.SubType = subfieldType;
                if (fieldValues.Length > 0)
                {
                    foreach (String s1 in fieldValues.Split('|'))
                    {
                        if ((s1.Length > 0))
                        {
                            String[] s2 = s1.Split('-');
                            oUDF.ValidValues.Description = s2[1];
                            oUDF.ValidValues.Value = s2[1];
                            oUDF.ValidValues.Add();
                        }

                    }
                }
                oUDF.DefaultValue = defaultValue;
                oUDF.LinkedTable = linkTable;
                if (oUDF.Add() == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                    oUDF = null;
                    GC.Collect();

                    return "Successfully Added UDF";
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                    oUDF = null;
                    GC.Collect();
                    string dat = sapCompany.GetLastErrorDescription();
                    lastErrorMessage += Environment.NewLine + DBCompany + "- Error Adding UDF : " + sapCompany.GetLastErrorDescription();

                    return "Error";
                }

            }
            catch (Exception ex)
            {
                return ex.Message;
                throw ex;
            }

        }

        public bool createUDT(String tableName, String description, SAPbobsCOM.BoUTBTableType tableType)
        {

            try
            {
                int iRet = -1;
                SAPbobsCOM.UserTablesMD ouTables = default(SAPbobsCOM.UserTablesMD);

               ouTables = (SAPbobsCOM.UserTablesMD)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!ouTables.GetByKey(tableName))
                {
                    ouTables.TableName = tableName;
                    ouTables.TableDescription = description;
                    ouTables.TableType = tableType;
                    iRet = ouTables.Add();
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(ouTables)
                    // ouTables = Nothing
                }
                //GC.Collect()
                if (iRet == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ouTables);
                    ouTables = null;
                    GC.Collect();
                    return true;
                }
                else
                {
                    lastErrorMessage += "Fail to Add UDT " + sapCompany.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ouTables);
                    ouTables = null;
                    GC.Collect();
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public bool CreateSettings()
        {
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    createUDT("AF_COMP", "Company Mapping Table", BoUTBTableType.bott_NoObject);
                    createUDF("@AF_COMP", "AF_SAPCOMPANY", "SAP COMPANY", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("@AF_COMP", "AF_MBSCOMP", "MBS COMPANY", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);

                    createUDF("OPCH", "AF_INVNO", "Invoice Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
                    createUDF("OPCH", "AF_PONO", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
                    createUDF("OPCH", "AF_PODESC", "PO Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("OPCH", "AF_ISPAY", "AF_ISPAID", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("OPCH", "AF_INVTYPE", "INVOICE TYPE", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("OPCH", "AF_INVEXPORT", "Exported", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("PCH1", "AF_LINENO", "AF LINENO", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);



                }
                else
                {
                    
                    return false;
                }
                return true;
                //createUDF("OPCH", "FCURRAMT", "Foreign Currency Amount", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Price)
                      
            }
            catch (Exception ex)
            {
                lastErrorMessage += Environment.NewLine + "Exception Error : " + ex.Message;
                return false;
            }
        }

        public bool CreateSettingsFORMagento()
        {
            try
            {
                if (connectToLoginCompany(ConstantClass.SBOServer, ConstantClass.Database, ConstantClass.SQLUserName, ConstantClass.SQLPassword, ConstantClass.SAPUser, ConstantClass.SAPPassword))
                {
                    createUDT("AF_COMP", "Company Mapping Table", BoUTBTableType.bott_NoObject);
                    createUDF("@AF_COMP", "AF_SAPCOMPANY", "SAP COMPANY", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("@AF_COMP", "AF_MBSCOMP", "MBS COMPANY", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    //Webcon Authenticat Config
                    createUDT("CXA_WEBCON", "WEBCON Authentication Config", BoUTBTableType.bott_NoObject);
                    createUDF("@CXA_WEBCON", "Password", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("@CXA_WEBCON", "DateEnd", "Date End", SAPbobsCOM.BoFieldTypes.db_Date, 11);

                    //Data UDF's HEADER
                    createUDF("ORDR", "AM_ORDERID", "Magento Order ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                    createUDF("ORDR", "AM_ORDERSTATUS", "Magento Order Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("ORDR", "AM_CARTBEFDISC", "Magento Cart Before Discount", SAPbobsCOM.BoFieldTypes.db_Float, 254, BoFldSubTypes.st_Price);
                    createUDF("ORDR", "AM_GRANDTOTAL", "Magento Grand Total", SAPbobsCOM.BoFieldTypes.db_Float, 254, BoFldSubTypes.st_Price);
                    createUDF("ORDR", "AM_EMAIL", "Magento Buyer's Email", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("ORDR", "AM_FULLNAME", "Magento Full Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("ORDR", "AM_PURCHASEDATE", "Magento Purchase Date", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                    createUDF("ORDR", "AM_COMPLETIONDATE", "Magento Completion Date", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                    createUDF("ORDR", "AM_MOP", "Magento MOP", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("ORDR", "AM_STRIPE", "Magento Stripe", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("ORDR", "AM_STRIPESTATUS", "Magento Stripe", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    //Data UDF Details
                    createUDF("RDR1", "AM_ORDERID", "Magento Order ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_TYPE", "Magento Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_STATUS", "Magento Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_SUBORDERID", "Magento Sub Order Id", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_VOUCHERID", "Magento VoucherID", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    //Combination of SUBORDERID and VoucherId.
                    createUDF("RDR1", "AM_PRODID", "Magento Prod ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_PRODNAME", "Magento Prod Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_UNIQUEKEY", "Magento UNIQUEKEY", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
                    createUDF("RDR1", "AM_QUANTITY", "Magento Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 254, BoFldSubTypes.st_Quantity);
                    createUDF("RDR1", "AM_PROPRICEGST", "Magento Provider Price GST", SAPbobsCOM.BoFieldTypes.db_Float, 11, BoFldSubTypes.st_Price);
                    createUDF("RDR1", "AM_COMMISSIONRATE", "Magento Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, 11, BoFldSubTypes.st_Price);
                    //
                    createUDF("RDR1", "AM_COMMAMTEXGST", "Magento Commission Amt w/o GST", SAPbobsCOM.BoFieldTypes.db_Float, 11, BoFldSubTypes.st_Price);
                    createUDF("RDR1", "AM_PRODUCTDISCOUNT", "Magento Product Discount", SAPbobsCOM.BoFieldTypes.db_Float, 11, BoFldSubTypes.st_Price);
                    createUDF("RDR1", "AM_GSTTYPE", "Magento GST Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);                 
                    createUDF("RDR1", "MRV_NMRV", "Magento MRV_NMRV", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                    createUDF("RDR1", "PAYMENTTYPE", "Magento Payment Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 150);
                    //UNIQUEKEY
                    //


                    

                }
                else
                {

                    return false;
                }
                return true;
                //createUDF("OPCH", "FCURRAMT", "Foreign Currency Amount", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Price)

            }
            catch (Exception ex)
            {
                lastErrorMessage += Environment.NewLine + "Exception Error : " + ex.Message;
                return false;
            }
        }
        #endregion
    
        #region " Help "
        string GetCostCenter(string CostCenterCode)
        {
            try
            {
                CostCode = CostCenterCode;
                return CostCenterCode;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        string GetVATGROUP(string VatCode)
        {
            try
            {
                VATCode = VatCode;
                return VatCode;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        string GETGLACCOUNT(string GLACCOUNT)
        {
            try
            {
                GLAccount = GLACCOUNT;
                return GLACCOUNT;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string DataString(string value, int length)
        {
            if (String.IsNullOrEmpty(value)) return string.Empty;

            return value.Length <= length ? value : value.Substring(value.Length - length);
        }
           #endregion

        #region " Properties "
        public string LastErrorMessage { 
            get
            {
                return lastErrorMessage;
            } 
        }
        public SAPbobsCOM.Company SAPCOMPANYOBJECT
        {
            get
            {
                return sapCompany;
            }
            set
            {
                SAPCOMPANYOBJECT = sapCompany;
            }
        }
        #endregion
    }
    public class ConstantClass
    {

        public static string SBOServer = "";
        public static string SQLUserName = "";
        public static string SQLPassword = "";
        public static int SQLVersion = 2014;
        public static string SAPUser = "";
        public static string SAPPassword = "";
        public static string Database = "";
        //=======================================================
        //Service provided by Telerik (www.telerik.com)
        //Conversion powered by NRefactory.
        //Twitter: @telerik
        //Facebook: facebook.com/telerik
        //=======================================================

    }
    public class CompanyData
    {
        public string CompanyCode { get; set; }
        public string CompanyName { get; set; }
    }
}
