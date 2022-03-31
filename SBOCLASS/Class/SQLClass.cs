using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
namespace SBOCLASS.Class
{
    public class SQLClass : Interface.ISQL
    {
        string strConnString;
        CommandType cmdType;

        private DataTable _getDataTable(string sqlQuery)
        {
            try
            {
                DataTable dt = null;
                using (SqlConnection cn = new SqlConnection(strConnString))
                {
                    SqlCommand cmd = cn.CreateCommand();
                    string sql = string.Empty;
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.Text;
                    //cmd.Parameters.Add("@Status", SqlDbType.VarChar).Value = Status
                    cmd.CommandText = sqlQuery;

                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        using (DataSet ds = new DataSet())
                        {
                            da.Fill(ds);
                            dt = ds.Tables[0];
                        }
                    }
                    return dt;
                }
            }
            catch
            {
                throw;
            }
        }
        //Execute Commands
        private bool _executeCommand(string sqlQuery)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(strConnString))
                {

                    cn.Open();
                    SqlCommand cmd = cn.CreateCommand();
                    string sql = string.Empty;
                    SqlTransaction trans = cn.BeginTransaction();

                    try
                    {
                        sql = sqlQuery;
                        cmd.Connection = cn;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = sql;
                        cmd.Transaction = trans;
                        //Executes Commands
                        cmd.ExecuteNonQuery();

                        trans.Commit();
                        trans = null;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        throw ex;
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        private object _executeScalar(string sqlQuery)
        {
            try
            {
                object retObj = null;

                using (SqlConnection cn = new SqlConnection(strConnString))
                {
                    SqlCommand cmd = cn.CreateCommand();
                    string sql = string.Empty;
                    cmd.Connection = cn;
                    cmd.CommandType = cmdType;
                    //cmd.Parameters.Add("@Status", SqlDbType.VarChar).Value = Status
                    cmd.CommandText = sqlQuery;
                    cn.Open();
                    retObj = cmd.ExecuteScalar();
                    cn.Close();
                }
                return retObj;
            }
            catch
            {
                throw;
            }
        }

        public static bool UpdateExecSQL(string sqlTable, int DocEntry, string ConnectionString)
        {
            try
            {
                //(System.Configuration.ConfigurationManager.ConnectionStrings["connectionStringName"].ConnectionString
                using (SqlConnection cn = new SqlConnection(ConnectionString))
                {

                    cn.Open();
                    SqlCommand cmd = cn.CreateCommand();
                    string sql = string.Empty;
                    SqlTransaction trans = cn.BeginTransaction();

                    try
                    {
                        sql = "Update " + sqlTable + " set U_IsCheck = '1' where docentry = @DocEntry";
                        cmd.Connection = cn;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = sql;
                        cmd.Transaction = trans;

                        //  cmd.Parameters.Add("@sqlTable", SqlDbType.).Value = _sqlTable
                        cmd.Parameters.Add("@DocEntry", SqlDbType.VarChar).Value = DocEntry;
                        cmd.ExecuteNonQuery();

                        trans.Commit();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        throw ex;
                    }

                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        
        #region " Properties "
        public string ConnectionString
        {
            get
            {
                return strConnString;
            }
            set
            {
                strConnString = value;
            }
        }

        public System.Data.DataTable GetDataTable(string sqlQuery)
        {
            return _getDataTable(sqlQuery);
        }

        public bool ExecuteCommand(string sqlQuery)
        {
            throw new NotImplementedException();
        }

        public System.Data.CommandType CommandType
        {
            get
            {
                return cmdType;
            }
            set
            {
                cmdType = value;
            }
        }

        public object ExecuteScalar(string sqlQuery)
        {
            return _executeScalar(sqlQuery);
        }
        #endregion

    }
}
