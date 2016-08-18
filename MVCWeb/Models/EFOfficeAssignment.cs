using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class EFOfficeAssignment
    {
        [Key]
        [ForeignKey("EFInstructor")]
        public int EFInstructorID { get; set; }
        [StringLength(50),Display(Name="Office Location")]
        public string Location { get; set; }
        public virtual EFInstructor EFInstructor { get; set; }
    }
}