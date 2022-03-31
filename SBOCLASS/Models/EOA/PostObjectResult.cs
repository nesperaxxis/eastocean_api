using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOCLASS.Models.EOA
{
    public class PostObjectResult
    {
        public bool Status;
        public string Remark;
        public int DocEntry;
        public int DocNumber;
        public string ObjType;

        public PostObjectResult() { }

        public PostObjectResult(string err, bool success = false)
        {
            Status = success;
            Remark = err;
        }

        public PostObjectResult(string objType, int docEntry, int docNum)
        {
            Status = true;
            DocEntry = docEntry;
            DocNumber = docNum;
            ObjType = objType;
        }
    }
}
