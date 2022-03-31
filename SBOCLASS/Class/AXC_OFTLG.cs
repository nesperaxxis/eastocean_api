using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static SBOCLASS.Class.SBOSupport;

namespace SBOCLASS.Class
{
    class AXC_OFTLG
    {
        public const string OFTLG_TABLE_UID = "@AXC_OFTLG";
        public const string OFTLG_USER_ID = "U_AXC_OUSER";
        public const string OFTLG_WS_OBJECT_TYPE = "U_AXC_OBJTP";
        public const string OFTLG_WS_OBJECT_CODE = "U_AXC_OBJCD";
        public const string OFTLG_WS_OBJECT_NAME = "U_AXC_OBJNM";
        public const string OFTLF_WS_DIRECTION = "U_AXC_DRCTN";
        public const string OFTLG_WS_OPERATION = "U_AXC_OPRTN";
        public const string OFTLG_WS_POST_DATA = "U_AXC_PDATA";
        public const string OFTLG_WS_POST_RESULT = "U_AXC_PRSLT";
        public const string OFTLG_WS_POST_SUCCESS = "U_AXC_SCCES";
        public const string OFTLG_WS_EXPORT_TIME_STAMP = "U_AXC_TSTMP";
        public const string OFTLG_WS_EXTERNAL_KEY = "U_AXC_EXTID";

        public static void GenerateLogRecord(SAPbobsCOM.Company company, String ObjType, String ObjCode, String ObjName, String ExtenalID, Operation Ops, String Data, String Result, Boolean Success)
        {
            try
            {
                SAPbobsCOM.UserTable ut = company.UserTables.Item(OFTLG_TABLE_UID.Substring(1));

                ut.UserFields.Fields.Item(OFTLG_USER_ID).Value = company.UserSignature;
                ut.UserFields.Fields.Item(OFTLG_WS_OBJECT_TYPE).Value = ObjType ?? "";
                ut.UserFields.Fields.Item(OFTLG_WS_OBJECT_CODE).Value = ObjCode ?? "";
                ut.UserFields.Fields.Item(OFTLG_WS_OBJECT_NAME).Value = ObjName ?? "";
                ut.UserFields.Fields.Item(OFTLG_WS_OPERATION).Value = Enum.GetName(typeof(SBOSupport.Operation), Ops);
                ut.UserFields.Fields.Item(OFTLF_WS_DIRECTION).Value = "I";      //Inbound
                                                                                //if (Data.Length > 4000) Data = Data.Substring(0, 4000);
                ut.UserFields.Fields.Item(OFTLG_WS_POST_DATA).Value = Data;
                //if (Result.Length > 4000) Result = Result.Substring(0, 4000);
                ut.UserFields.Fields.Item(OFTLG_WS_POST_RESULT).Value = Result??"";
                ut.UserFields.Fields.Item(OFTLG_WS_POST_SUCCESS).Value = Success ? "Y" : "N";
                string currentTime = SBOSupport.GETSINGLEVALUE("SELECT CONVERT(NVARCHAR(25),GETDATE(),120)", company);
                ut.UserFields.Fields.Item(OFTLG_WS_EXPORT_TIME_STAMP).Value = currentTime;
                ut.UserFields.Fields.Item(OFTLG_WS_EXTERNAL_KEY).Value = ExtenalID ?? "";

                int err = ut.Add();
                if (err != 0)
                    throw new Exception($"Could not create log. {company.GetLastErrorDescription()}");

                SBOSupport.ReleaseComObject(ut);
            }
            catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
        }
    }
}
