using SelectHealth.Ops.AWD.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AWD_Create_Helper
{
    internal class AWD_User
    {
        public string Username { get; set; }
        public string Business_Area { get; set; }
        public string Work_Select { get; set; }
        public string Work_Group { get; set; }
        public string Security_Level { get; set; }
        public string Redirect { get; set; }
        public string Country { get; set; }
        public string In_By { get; set; }
        public string Out_By { get; set; }
        public string Phone { get; set; }
        public AccountService.Status Status { get; set; }
        public string Forward_Queue { get; set; }
        public string Work_Action { get; set; }
        public bool Personal_Queue { get; set; }
        public string Alias { get; set; }
        public string Last_Name { get; set; }
        public string Middle_Name { get; set; }
        public string First_Name { get; set; }
    }
}
