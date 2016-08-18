using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public enum Grade
    { 
        A,B,C,D,F
    }
    public class EFEnrollment
    {
        public int EFEnrollmentID { get; set; }
        //EF interprets a property as a foreign key propery if it's named <navigation property name><primary key property name>
        public int EFCourseID { get; set; }
        public int EFStudentID { get; set; }
        //? indicates that the Grade is nullable, a grade that is null is different from a zeor grade,
        //null means a grade is not known or has not been assigned
        [DisplayFormat(NullDisplayText="No grade")]
        public Grade? Grade { get; set; }

        public virtual EFCourse EFCourse { get; set; }
        public virtual EFStudent EFStudent { get; set; }
    }
}