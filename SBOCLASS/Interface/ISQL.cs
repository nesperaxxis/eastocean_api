using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
namespace SBOCLASS.Interface
{
    public interface ISQL
    {
        string ConnectionString { get; set; }
        System.Data.DataTable GetDataTable(string sqlQuery);
        Boolean ExecuteCommand(string sqlQuery );
        CommandType CommandType { get; set; }
        Object ExecuteScalar(string sqlQuery );   

    }
}
