using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class EFDepartment
    {
        public int EFDepartmentID { get; set; }
        [StringLength(50,MinimumLength=3)]
        public string Name { get; set; }
        [DataType(DataType.Currency),Column(TypeName="money")]
        public decimal Budget { get; set; }
        [DataType(DataType.Date),DisplayFormat(DataFormatString="{0:yyyy-MM-dd}",ApplyFormatInEditMode=true)]
        [Display(Name="Start Date")]
        public DateTime StartDate { get; set; }
        public int? EFInstructorID { get; set; }
        public virtual EFInstructor Administrator { get; set; }
        public virtual ICollection<EFCourse> EFCourses { get; set; }
    }
}