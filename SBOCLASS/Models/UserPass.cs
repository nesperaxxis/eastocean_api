using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;

namespace SBOCLASS.Models
{
    public class UserPass
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public DateTime EndDate { get; set; }
    }
}
