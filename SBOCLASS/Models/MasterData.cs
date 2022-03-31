using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;


namespace SBOCLASS.Models
{
    public class ItemGL
    {
        public string ItemID { get; set; }
        public string GLCode { get; set; }
    }
    public class Item
    {
        public string ItemID { get; set; }
        public string Name { get; set; }
        public string UOMID { get; set; }
    }

    public class Department
    {
        public string DepartmentID { get; set; }
        public string Name { get; set; }

    }
    public class GLACCT
    {
        public string GLCode { get; set; }
        public string GLName { get; set; }

    }
    public class GLBudget
    {
        public int Year { get; set; }
        public string GLCode { get; set; }
        public string BudgetScenario { get; set; }
        public string Department { get; set; }
        public string GLBudgetVal { get; set; }
    }
    public class Project
    {
        public string ProjectID { get; set; }
        public string Name { get; set; }
    }
    public class Supplier
    {
        public string SupplierID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }
    public class SupplierContact
    {
        public string SupplierID { get; set; }
        public string ContactID { get; set; }
        public string ContactPerson { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
    }
}
