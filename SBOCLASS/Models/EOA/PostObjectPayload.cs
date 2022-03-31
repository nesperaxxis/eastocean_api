using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models.EOA
{
    public class PostObjectPayload
    {
        public const string SYNCH_O_OBJECT_VENDOR = "2V";
        public const string SYNCH_O_OBJECT_CUSTOMER = "2C";
        public const string SYNCH_O_OBJECT_ITEM = "4";
        public const string SYNCH_O_OBJECT_ITEM_CATEGORY = "52";
        public const string SYNCH_O_OBJECT_BRAND = "43";
        public const string SYNCH_O_OBJECT_BAR_CODE = "1470000062";
        public const string SYNCH_O_OBJECT_WAREHOUSE = "64";
        public const string SYNCH_O_OBJECT_BIN = "10000206";
        public const string SYNCH_O_OBJECT_BOM = "66";
        public const string SYNCH_O_OBJECT_SALES_ORDER = "17";
        public const string SYNCH_O_OBJECT_RESERVE_INVOICE = "13R";
        public const string SYNCH_O_OBJECT_AR_CN = "14";
        public const string SYNCH_O_OBJECT_AR_RETURNS = "16";
        public const string SYNCH_O_OBJECT_PURCHASE_ORDER = "22";
        public const string SYNCH_O_OBJECT_AP_RETURN = "21";
        public const string SYNCH_O_OBJECT_AP_CN = "19";
        public const string SYNCH_O_OBJECT_WORK_ORDER = "202";
        public const string SYNCH_I_OBJECT_PICK_LIST = "156";
        public const string SYNCH_I_OBJECT_AR_RETURN = "16";
        public const string SYNCH_I_OBJECT_GRPO = "20";
        public const string SYNCH_I_OBJECT_TR_REQUEST = "1250000001";
        public const string SYNCH_I_OBJECT_WHS_TRANSFER = "67";
        public const string SYNCH_I_OBJECT_ISSUE_PROD = "60P";
        public const string SYNCH_I_OBJECT_RECPT_PROD = "59P";
        public const string SYNCH_I_OBJECT_STOCK_ADJ_NEG = "60";
        public const string SYNCH_I_OBJECT_STOCK_ADJ_POS = "59";
        public const string SYNCH_I_OBJECT_DELIVERY_ORDER = "15";
        public const string SYNCH_O_OBJECT_STOCK_COUNT = "1470000065";
        public const string SYNCH_I_OBJECT_STOCK_POST = "10000071";

        public string ObjType { get; set; }
        public string BaseObj { get; set; }
        public object Data { get; set; }


    }
}
