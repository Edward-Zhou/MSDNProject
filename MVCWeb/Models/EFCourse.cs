using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class EFCourse
    {
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.None)]
        [Display(Name="Number")]
        public int EFCourseID { get; set; }
        [StringLength(50,MinimumLength=3)]
        public string Title { get; set; }
        [Range(0,5)]
        public int Credits { get; set; }

        public int EFDepartmentID { get; set; }
        public virtual EFDepartment EFDepartment { get; set; }
        public virtual ICollection<EFInstructor> EFInstructors { get; set; }

        public virtual ICollection<EFEnrollment> Enrollments { get; set; }
    }
}