using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AXC_EOA_WMSWebAPI.Class
{
    public class ConstantClass
    {

        public static string SBOServer = System.Configuration.ConfigurationManager.AppSettings["SBOServer"].ToString();
        public static string SQLUserName = System.Configuration.ConfigurationManager.AppSettings["SQLUserName"].ToString();
        public static string SQLPassword = System.Configuration.ConfigurationManager.AppSettings["SQLPassword"].ToString();
        public static int SQLVersion = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SQLVersion"].ToString());
        public static string SAPUser = System.Configuration.ConfigurationManager.AppSettings["SAPUser"].ToString();
        public static string SAPPassword = System.Configuration.ConfigurationManager.AppSettings["SAPPassword"].ToString();
        public static string Database = System.Configuration.ConfigurationManager.AppSettings["Database"].ToString();
        //=======================================================
        //Service provided by Telerik (www.telerik.com)
        //Conversion powered by NRefactory.
        //Twitter: @telerik
        //Facebook: facebook.com/telerik
        //=======================================================

    }
}