using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class DefaultModel
    {
    }
        public class CustomerAccount
        {
            public int  ID { get; set; }
            public string Full_Name { get; set; }
            public string Email { get; set; }
            public bool Is_Biz_Cust { get; set; }
            
            [DataType(DataType.DateTime)]
            [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        
            public DateTime DOB { get; set; }
        }
        public class custdetail
        {
            public string Company_Name { get; set; }
            public string Company_Email { get; set; }            
        }
    
}