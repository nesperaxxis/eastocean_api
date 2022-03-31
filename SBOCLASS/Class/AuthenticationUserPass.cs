using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SBOCLASS.Interface;

namespace SBOCLASS.Class
{
    public class AuthenticationUserPass
    {
        public int AuthenticateUserPass(string SQLConnection,string UserName, string Password)
        {
            try
            {
                //Declarations
                int ReturnCount = 0;
                DateTime now = DateTime.Now;

                List<Models.UserPass> AuthModelData = new List<Models.UserPass>();
                          
                System.Data.DataTable dt;
                // System.Data.DataTable dtDetail;

                SQLClass sql = new SQLClass();
                sql.ConnectionString = SQLConnection; // System.Configuration.ConfigurationManager.ConnectionStrings["connectionStringName"].ConnectionString;
                
                sql.CommandType = System.Data.CommandType.Text;
                dt = sql.GetDataTable("SELECT * FROM [@CXA_WEBCON] WHERE Name = '"+ UserName +"' and U_Password='"+ Password +"'");
                if (dt.Rows.Count <= 0)
                {
                    ReturnCount = 1;//for invalid Username and Password
                }
                else
                {
                    AuthModelData = (from System.Data.DataRow row in dt.Rows

                                     select new Models.UserPass
                                     {
                                         Username = row[1].ToString(),
                                         Password = row[2].ToString(),
                                         EndDate = (DateTime)row[3],
                                     }).ToList();

                    var data = AuthModelData.Where(n => n.EndDate >= now ).Select(n => new { n.Username, n.Password }).ToList();
                    if (data.Count <= 0)
                    {
                        ReturnCount =  2;//for expired Username and Password
                    }
                }
                return ReturnCount; // valid Username and Password 

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
